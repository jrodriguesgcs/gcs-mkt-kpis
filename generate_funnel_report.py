#!/usr/bin/env python3
"""
HubSpot Traffic-Source Funnel report generator.

Fetches every contact created on/after a cutoff date (see the trade-off
in fetch_contacts()'s docstring; override with FETCH_CONTACTS_SINCE),
pre-filtered server-side by the four overall filters, + its associated
deals from HubSpot. The same four filters are re-applied client-side as
the authoritative pass. Writes a styled Excel workbook to
reports/funnel_report.xlsx with one sheet per funnel -- "New Contacts",
"New Deals", "Existing Deals":

  Rows are Original Traffic Source, 3 levels deep (Source -> Drill-Down 1
  -> Drill-Down 2). Columns are that funnel's own applicable Stages (a
  stage that doesn't apply to a given funnel is simply omitted from that
  funnel's sheet, not placeholder-filled) x Time (Month -> Week -> Day,
  year-to-date -- January through the current month, not the full
  calendar year). Both axes use Excel's native Group & Outline (+/-
  buttons) so a viewer can expand or collapse either hierarchy one level
  at a time without touching a pivot table -- every row and column starts
  fully collapsed to its coarsest level (Source; Month) when the file is
  opened.

Run with:  python generate_funnel_report.py

Deliberately NOT in scope here: GitHub Actions / cron scheduling, email
delivery, and secrets management beyond a local .env file. See README.md.
"""

from __future__ import annotations

import colorsys
import os
import random
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline

HUBSPOT_API_BASE = "https://api.hubapi.com"
OUTPUT_PATH = os.path.join("reports", "funnel_report.xlsx")

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# =========================================================================
# GCS Design System
# =========================================================================
NIGHT_BLUE = "000957"
ELECTRIC_BLUE = "3F8CFF"
SLATE = "414856"
BORDER_TINT = "E3EDFF"

FONT_BODY = "Heebo"
FONT_TITLE = "Yrsa"
FONT_MONO = "JetBrains Mono"

THIN_BORDER = Border(*(Side(style="thin", color=BORDER_TINT) for _ in range(4)))
HEADER_FILL = PatternFill(fill_type="solid", fgColor=NIGHT_BLUE)
BODY_FONT = Font(name=FONT_BODY, color=SLATE)
BOLD_BODY_FONT = Font(name=FONT_BODY, color=SLATE, bold=True)
MONO_FONT = Font(name=FONT_MONO, color=SLATE)
CORNER_TITLE_FONT = Font(name=FONT_TITLE, color="FFFFFF", size=12)  # Yrsa, weight 400 -- titles only

SEARCH_API_DELAY_SECONDS = 0.25
MAX_RETRIES = 5

# =========================================================================
# Colour helpers -- HSL ramp per funnel (hue/sat held constant, lightness
# interpolated 85% -> 25% across 6 steps), and per-cell text-colour choice.
# =========================================================================


def hex_to_rgb01(hex_color: str) -> tuple[float, float, float]:
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return r, g, b


def rgb01_to_hex(r: float, g: float, b: float) -> str:
    return "".join(f"{round(max(0.0, min(1.0, c)) * 255):02X}" for c in (r, g, b))


def hsl_ramp(base_hex: str, n: int = 6, l_from: float = 0.85, l_to: float = 0.25) -> list[str]:
    """n hex colours from base_hex, hue/saturation held constant, lightness
    interpolated linearly from l_from down to l_to (lightest first)."""
    r, g, b = hex_to_rgb01(base_hex)
    h, _l, s = colorsys.rgb_to_hls(r, g, b)
    ramp = []
    for i in range(n):
        t = i / (n - 1) if n > 1 else 0.0
        lightness = l_from + (l_to - l_from) * t
        rr, gg, bb = colorsys.hls_to_rgb(h, lightness, s)
        ramp.append(rgb01_to_hex(rr, gg, bb))
    return ramp


def readable_text_color(bg_hex: str) -> str:
    """White text below ~55% background lightness, else Night Blue."""
    r, g, b = hex_to_rgb01(bg_hex)
    _h, l, _s = colorsys.rgb_to_hls(r, g, b)
    return "FFFFFF" if l < 0.55 else NIGHT_BLUE


# =========================================================================
# Funnel / stage definitions
# =========================================================================
STAGE_LABELS = ["New Contacts", "New Deals", "Qualified Deals",
                "Opportunities", "Proposals Sent", "Proposal Signed"]

FUNNELS = [
    # All 3 sheets share the same colour pattern (the "New Deals"/Electric
    # Blue ramp) per request -- base_color is uniform, not per-funnel.
    {"name": "New Contacts", "base_color": ELECTRIC_BLUE,
     "applicable": [True, True, True, True, True, True]},
    {"name": "New Deals", "base_color": ELECTRIC_BLUE,
     "applicable": [False, True, True, True, True, True]},
    {"name": "Existing Deals", "base_color": ELECTRIC_BLUE,
     "applicable": [False, False, True, True, True, True]},
]

# Target labels the overall filters are defined against (values are
# resolved dynamically from each property's live enum options -- never
# hardcoded, per the spec's "search properties by label text" instruction).
CONTACT_TYPE_EXCLUDE_LABELS = ["B2B Partnership Development", "B2B Institutional Relations"]
LEAD_SOURCE_EXCLUDE_LABELS = ["Bundle Offer", "Other", "Instantly", "Private", "Walk-In",
                              "Email", "Partner Referral", "Phone Calls", "Events", "Client Referral"]
BRAND_REQUIRED_LABEL = "Global Citizen Solutions"
BRAND_DOMAIN_EXCLUDE_LABELS = ["BePortugal"]

# hs_all_assigned_business_unit_ids (the "Brands" property backing HubSpot's
# Business Units feature) is declared with "externalOptions": true on
# GET /crm/v3/properties/contacts/hs_all_assigned_business_unit_ids -- this
# is HubSpot's signal that the property's valid option list is NOT embedded
# in the Properties API response (its "options" array is always []); the
# real id<->label mapping lives in a separate Business Units API instead.
# That's true for every credential (Service Key, Private App, admin --
# confirmed against the live portal), not a permissions quirk of any one
# token. The dedicated Business Units API
# (GET /business-units/v3/business-units/user/{userId}) would be the
# "proper" way to resolve this, but it needs the business_units_view.read
# scope (not granted here) and is a per-user lookup with no account-wide
# "list all brands" endpoint -- not a good fit for a script that otherwise
# resolves everything dynamically. So this one mapping is hardcoded,
# sourced directly from the portal's own Business Units search-options UI
# (Label / Internal name / value, confirmed live): if a business unit is
# ever renamed or a new one added, resolve_reference_data() below still
# raises HubSpotError rather than silently using a stale id.
BUSINESS_UNIT_LABEL_TO_ID = {
    "Global Citizen Solutions": "0",
    "Get NIF Portugal": "18346255",
    "Goldcrest": "18395897",
}


# =========================================================================
# Date / calendar helpers
# =========================================================================


def month_start(d: date) -> date:
    return d.replace(day=1)


def month_end_exclusive(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def week_start(d: date) -> date:
    """Monday of the week containing d."""
    return d - timedelta(days=d.weekday())


def _has_brand_value(raw_value: str | None, required_value: str) -> bool:
    """True if required_value is present in raw_value, treating raw_value
    as a ';'-separated token list. The resolved Brand property may be a
    genuinely single-valued enum (a plain '==' would suffice) or HubSpot's
    multi-valued Business Units property (a contact can belong to more
    than one) -- token containment is correct for both without needing to
    know which case applies."""
    if not raw_value:
        return False
    return required_value in [token.strip() for token in raw_value.split(";")]


def parse_hs_datetime(value) -> datetime | None:
    """Parse a HubSpot ISO-8601 datetime (or date-only) string to a UTC
    datetime. HubSpot v3 properties come back as strings like
    '2024-05-01T00:00:00Z' or '2024-05-01T10:23:45.678Z'; some date-only
    properties come back as 'YYYY-MM-DD'."""
    if not value:
        return None
    v = str(value).strip()
    if v.endswith("Z"):
        v = v[:-1]
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(v, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


# =========================================================================
# Progress printing -- long-running steps (contact/deal fetch, workbook
# build) overwrite a single line so the run never looks stalled without
# spamming the console with one line per page/chunk.
# =========================================================================


def _print_progress(label: str, count: int, total: int | None = None) -> None:
    suffix = f"/{total:,}" if total is not None else ""
    # \x1b[K clears the rest of the line after the cursor -- without it, a
    # shorter new line (e.g. label text shrinking) would leave stray
    # characters from the previous, longer line visible past it.
    print(f"\r  ...{count:,}{suffix} {label}\x1b[K", end="", flush=True)


def _print_progress_done() -> None:
    print()  # move off the overwritten progress line


# =========================================================================
# HubSpot client (Bearer auth, retry/backoff, pagination helpers)
# =========================================================================


class HubSpotError(RuntimeError):
    """Raised when a HubSpot API call fails, or reference data can't be resolved."""


class HubSpotClient:
    """Thin wrapper around requests.Session with auth + backoff baked in."""

    def __init__(self, access_token: str):
        if not access_token:
            raise HubSpotError(
                "HUBSPOT_ACCESS_TOKEN is not set. Copy .env.example to .env "
                "and fill in a HubSpot Service Key or private app token."
            )
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        })

    def request(self, method: str, path: str, is_search: bool = False, **kwargs):
        if is_search:
            time.sleep(SEARCH_API_DELAY_SECONDS)
        url = f"{HUBSPOT_API_BASE}{path}"
        last_exc = None
        for attempt in range(MAX_RETRIES + 1):
            try:
                resp = self.session.request(method, url, timeout=30, **kwargs)
            except requests.RequestException as exc:
                last_exc = exc
                if attempt == MAX_RETRIES:
                    break
                time.sleep((2 ** attempt) + random.uniform(0, 1))
                continue

            if resp.status_code == 429:
                if attempt == MAX_RETRIES:
                    raise HubSpotError(
                        f"HubSpot rate limit exceeded after {MAX_RETRIES} retries on {method} {path}"
                    )
                retry_after = resp.headers.get("Retry-After")
                time.sleep(float(retry_after) if retry_after else (2 ** attempt) + random.uniform(0, 1))
                continue

            if resp.status_code >= 400:
                raise HubSpotError(f"HubSpot API error {resp.status_code} on {method} {path}: {resp.text[:500]}")

            return resp.json() if resp.text else {}

        raise HubSpotError(f"HubSpot request failed on {method} {path}: {last_exc}")

    def get(self, path: str, is_search: bool = False, **kwargs):
        return self.request("GET", path, is_search=is_search, **kwargs)

    def post(self, path: str, is_search: bool = False, **kwargs):
        return self.request("POST", path, is_search=is_search, **kwargs)

    def paginate(self, path: str, params: dict | None = None, results_key: str = "results",
                 label: str | None = None) -> list:
        params = dict(params or {})
        results = []
        while True:
            data = self.get(path, params=params)
            results.extend(data.get(results_key, []))
            if label:
                _print_progress(f"{label} fetched so far", len(results))
            next_after = data.get("paging", {}).get("next", {}).get("after")
            if not next_after:
                break
            params["after"] = next_after
        if label:
            _print_progress_done()
        return results

    def search_all(self, object_type: str, body: dict, label: str | None = None) -> list:
        """POST-paginate /crm/v3/objects/{type}/search using the `after`
        cursor.

        HubSpot's Search API hard-caps pagination at 10,000 results: once
        `after` would reach offset 10000, the *next* page request itself
        fails with a 400 ("There was a problem with the request") rather
        than returning an empty page -- confirmed against a real 400 log
        with `"after": "10000"` in the request body. So we stop and warn
        as soon as we've collected 10,000, instead of making that request
        and crashing. Narrow the search filters (a later report-year
        cutoff, or the overall filters below) to stay under this if a
        truncation warning shows up."""
        body = dict(body)
        body.setdefault("limit", 100)
        results = []
        first_page = True
        while True:
            data = self.post(f"/crm/v3/objects/{object_type}/search", is_search=True, json=body)
            results.extend(data.get("results", []))
            if label and first_page:
                # Diagnostic: HubSpot's own reported match count for these
                # filters, independent of how many pages we go on to fetch.
                # If this differs wildly from the final fetched count for
                # reasons other than the 10k cap below, that points at the
                # credential's data visibility, not the filter logic.
                reported_total = data.get("total")
                if reported_total is not None:
                    print(f"\n  HubSpot reports {reported_total:,} {label} total match these filters.")
                first_page = False
            if label:
                _print_progress(f"{label} fetched so far", len(results))
            if len(results) >= 10000:
                total = data.get("total")
                if total is not None and total > len(results):
                    print(f"\n  WARNING: HubSpot Search API's 10,000-result pagination cap was hit "
                          f"({total:,} total match the filter) -- {total - len(results):,} are missing "
                          f"from this fetch. Narrow the search filters to stay under the cap.")
                break
            next_after = data.get("paging", {}).get("next", {}).get("after")
            if not next_after:
                break
            body["after"] = next_after
        if label:
            _print_progress_done()
        return results

    def batch_read(self, object_type: str, ids: list[str], properties: list[str],
                    label: str | None = None) -> list:
        results = []
        ids = list(dict.fromkeys(ids))
        for i in range(0, len(ids), 100):
            chunk = ids[i:i + 100]
            body = {"properties": properties, "inputs": [{"id": obj_id} for obj_id in chunk]}
            data = self.post(f"/crm/v3/objects/{object_type}/batch/read", json=body)
            results.extend(data.get("results", []))
            if label:
                _print_progress(f"{label} processed", min(i + 100, len(ids)), total=len(ids))
        if label:
            _print_progress_done()
        return results

    def batch_read_associations(self, from_type: str, to_type: str, ids: list[str],
                                 label: str | None = None) -> dict:
        result_map: dict = {}
        ids = list(dict.fromkeys(ids))
        for i in range(0, len(ids), 100):
            chunk = ids[i:i + 100]
            body = {"inputs": [{"id": obj_id} for obj_id in chunk]}
            data = self.post(f"/crm/v4/associations/{from_type}/{to_type}/batch/read", json=body)
            for entry in data.get("results", []):
                from_id = str(entry.get("from", {}).get("id"))
                to_ids = [str(to.get("toObjectId")) for to in entry.get("to", [])]
                result_map[from_id] = to_ids
            if label:
                _print_progress(f"{label} processed", min(i + 100, len(ids)), total=len(ids))
        if label:
            _print_progress_done()
        return result_map


# =========================================================================
# Step 0: reference-data resolution
# =========================================================================


@dataclass
class ReferenceData:
    source_prop: str
    dd1_prop: str
    dd2_prop: str
    contact_type_prop: str
    contact_type_exclude_values: dict
    brand_prop: str
    brand_required_value: str
    brand_domain_prop: str
    brand_domain_exclude_values: dict
    lead_source_prop: str
    lead_source_exclude_values: dict
    sales_pipeline_id: str
    sales_pipeline_label: str
    sql_lost_prop: str
    owner_assigneddate_prop: str
    closedlost_stage_id: str
    closedlost_date_entered_prop: str
    proposal_sent_prop: str
    proposal_signed_prop: str
    flags: list = field(default_factory=list)  # human-readable discrepancy notes


def _by_label(props: list[dict], label: str) -> dict | None:
    target = label.strip().lower()
    for p in props:
        if (p.get("label") or "").strip().lower() == target:
            return p
    return None


def _by_label_contains(props: list[dict], *needles: str) -> dict | None:
    needles_l = [n.lower() for n in needles]
    for p in props:
        lbl = (p.get("label") or "").lower()
        if all(n in lbl for n in needles_l):
            return p
    return None


def _match_option_values(prop: dict, target_labels: list[str], flags: list) -> dict:
    """Map each target label to its enum option value, warning (not
    raising) on any target label the property's live options no longer
    contain -- the spec asks us to flag drift, not silently proceed as if
    nothing changed."""
    options_by_label = {(o.get("label") or "").strip().lower(): o.get("value")
                         for o in prop.get("options", [])}
    resolved = {}
    for label in target_labels:
        value = options_by_label.get(label.strip().lower())
        if value is None:
            msg = (f"WARNING: property '{prop['name']}' no longer has an option labeled "
                   f"'{label}' -- this exclusion will not match anything for that label.")
            print(f"  {msg}")
            flags.append(msg)
        else:
            resolved[label] = value
    return resolved


def resolve_reference_data(client: HubSpotClient) -> ReferenceData:
    print("=== Step 0: resolving reference data ===")
    flags: list = []

    contact_props = client.get("/crm/v3/properties/contacts").get("results", [])
    deal_props = client.get("/crm/v3/properties/deals").get("results", [])
    contact_props_by_name = {p["name"]: p for p in contact_props}
    deal_props_by_name = {p["name"]: p for p in deal_props}

    # --- Original Traffic Source + Drill-Downs 1/2
    source_prop = contact_props_by_name.get("hs_analytics_source")
    dd1_prop = contact_props_by_name.get("hs_analytics_source_data_1")
    dd2_prop = contact_props_by_name.get("hs_analytics_source_data_2")
    if not (source_prop and "original traffic source" in (source_prop.get("label") or "").lower()):
        source_prop = _by_label_contains(contact_props, "original traffic source") or source_prop
    if not (dd1_prop and "drill-down 1" in (dd1_prop.get("label") or "").lower()):
        dd1_prop = _by_label_contains(contact_props, "original traffic source", "drill-down 1") or dd1_prop
    if not (dd2_prop and "drill-down 2" in (dd2_prop.get("label") or "").lower()):
        dd2_prop = _by_label_contains(contact_props, "original traffic source", "drill-down 2") or dd2_prop
    if not (source_prop and dd1_prop and dd2_prop):
        raise HubSpotError("Could not resolve Original Traffic Source / Drill-Down 1 / Drill-Down 2 contact properties.")
    print(f"  Original Traffic Source            = {source_prop['name']} (label: {source_prop.get('label')})")
    print(f"  Original Traffic Source Drill-Down 1 = {dd1_prop['name']} (label: {dd1_prop.get('label')})")
    print(f"  Original Traffic Source Drill-Down 2 = {dd2_prop['name']} (label: {dd2_prop.get('label')})")

    # --- Contact Type
    contact_type_prop = contact_props_by_name.get("contact_type") or _by_label(contact_props, "Contact Type")
    if not contact_type_prop:
        raise HubSpotError("Could not resolve the 'Contact Type' contact property.")
    contact_type_exclude_values = _match_option_values(contact_type_prop, CONTACT_TYPE_EXCLUDE_LABELS, flags)
    if len(contact_type_exclude_values) < len(CONTACT_TYPE_EXCLUDE_LABELS):
        raise HubSpotError("Contact Type is missing one of the required exclusion values; aborting.")
    print(f"  Contact Type                        = {contact_type_prop['name']} "
          f"(excludes: {list(contact_type_exclude_values.values())})")

    # --- Brand (no property is guaranteed to be labeled literally "Brand")
    #
    # More than one property can plausibly be "Brand": an exact-label match,
    # HubSpot's own multi-valued Business Units property
    # (hs_all_assigned_business_unit_ids, label "Brands" -- present and
    # populated on essentially every contact when Business Units is
    # enabled), and any other enum property whose options happen to include
    # "Global Citizen Solutions". A first pass at this (label-substring-only
    # tie-break) picked a property called "Client x Brand" that turned out
    # to be set on ~0.06% of contacts -- clearly a manual post-sale tag, not
    # a general "which brand" field -- while hs_all_assigned_business_unit_ids
    # covers ~92.6% of the portal, confirmed against live data and the
    # user's explicit choice between the two. So: prefer an exact-label
    # "Brand" match first, then hs_all_assigned_business_unit_ids
    # specifically (by name, not just by label substring) if it carries the
    # required option, and only fall back to the old generic scan -- which
    # is one property-population check away from repeating the same
    # mistake -- as a last resort.
    #
    # IMPORTANT: hs_all_assigned_business_unit_ids is declared with
    # "externalOptions": true -- its "options" array from the Properties
    # API is ALWAYS empty (confirmed live, GET
    # /crm/v3/properties/contacts/hs_all_assigned_business_unit_ids),
    # regardless of credential. So it can never be matched via the normal
    # "does this property's options contain the required label" check used
    # for every other property below -- if we required that, we'd always
    # fall through to the wrong "Client x Brand" property. It's special-
    # cased: once resolved by name, its required value comes from the
    # hardcoded BUSINESS_UNIT_LABEL_TO_ID map (see its definition for why).
    brand_prop = _by_label(contact_props, "Brand")
    if brand_prop is None:
        brand_prop = contact_props_by_name.get("hs_all_assigned_business_unit_ids")
    business_unit_brand = brand_prop is not None and brand_prop["name"] == "hs_all_assigned_business_unit_ids"
    if brand_prop is None:
        candidates = [p for p in contact_props if p.get("type") == "enumeration"
                      and any((o.get("label") or "").strip() == BRAND_REQUIRED_LABEL for o in p.get("options", []))]
        if not candidates:
            raise HubSpotError(f"No contact property has an enum option '{BRAND_REQUIRED_LABEL}'.")
        brand_prop = next((p for p in candidates if "brand" in (p.get("label") or "").lower()), candidates[0])
    if brand_prop.get("label", "").strip().lower() != "brand":
        msg = (f"NOTE: no contact property is labeled exactly 'Brand'; using '{brand_prop['name']}' "
               f"(label '{brand_prop.get('label')}') -- see the comment above resolve_reference_data's "
               f"Brand section for why this one and not another 'brand-ish' candidate.")
        print(f"  {msg}")
        flags.append(msg)
    if business_unit_brand:
        required_value = BUSINESS_UNIT_LABEL_TO_ID.get(BRAND_REQUIRED_LABEL)
        if required_value is None:
            raise HubSpotError(
                f"BUSINESS_UNIT_LABEL_TO_ID has no entry for '{BRAND_REQUIRED_LABEL}' -- a business unit was "
                f"likely renamed or removed; update the hardcoded map (see its definition for why it's hardcoded)."
            )
        brand_values = {BRAND_REQUIRED_LABEL: required_value}
    else:
        brand_values = _match_option_values(brand_prop, [BRAND_REQUIRED_LABEL], flags)
        if BRAND_REQUIRED_LABEL not in brand_values:
            raise HubSpotError(f"Resolved Brand property '{brand_prop['name']}' has no option '{BRAND_REQUIRED_LABEL}'.")
    print(f"  Brand                                = {brand_prop['name']} (label: {brand_prop.get('label')}), "
          f"required value = {brand_values[BRAND_REQUIRED_LABEL]!r} "
          f"(matched as a ';'-separated token, since this property may be multi-valued)")

    # --- Brand Domain
    brand_domain_prop = _by_label(contact_props, "Brand Domain") or _by_label(contact_props, "Brand domain")
    if not brand_domain_prop:
        raise HubSpotError("Could not resolve the 'Brand Domain' contact property.")
    brand_domain_exclude_values = _match_option_values(brand_domain_prop, BRAND_DOMAIN_EXCLUDE_LABELS, flags)
    if not brand_domain_exclude_values:
        raise HubSpotError("Brand Domain has none of its required exclusion values; aborting.")
    print(f"  Brand Domain                         = {brand_domain_prop['name']} "
          f"(excludes: {list(brand_domain_exclude_values.values())})")

    # --- Lead Source (spec pins this to the literal 'lead_source' property)
    lead_source_prop = contact_props_by_name.get("lead_source")
    if not lead_source_prop:
        msg = "WARNING: contact property 'lead_source' no longer exists on this portal."
        print(f"  {msg}")
        flags.append(msg)
        raise HubSpotError("Could not resolve the 'lead_source' contact property (see WARNING above).")
    lead_source_exclude_values = _match_option_values(lead_source_prop, LEAD_SOURCE_EXCLUDE_LABELS, flags)
    print(f"  Lead Source                          = {lead_source_prop['name']} "
          f"(excludes: {list(lead_source_exclude_values.values())})")

    # --- Sales pipeline + Closed Lost stage
    pipelines = client.get("/crm/v3/pipelines/deals").get("results", [])
    sales_pipeline = next((p for p in pipelines if p.get("id") == "default"), None)
    if not sales_pipeline:
        sales_pipeline = next((p for p in pipelines if "sales pipeline" in (p.get("label") or "").lower()), None)
        if sales_pipeline:
            msg = (f"NOTE: pipeline id is not literally 'default'; using '{sales_pipeline['id']}' "
                   f"(label '{sales_pipeline.get('label')}') matched by 'sales pipeline' in its label.")
            print(f"  {msg}")
            flags.append(msg)
    if not sales_pipeline:
        raise HubSpotError("Could not resolve the Sales pipeline (id 'default' or label containing 'sales pipeline').")
    print(f"  Sales pipeline                       = {sales_pipeline['id']} (label: {sales_pipeline.get('label')})")

    closedlost_stage = next(
        (s for s in sales_pipeline.get("stages", []) if "closed lost" in (s.get("label") or "").lower()), None
    )
    if not closedlost_stage:
        raise HubSpotError("Could not resolve the 'Closed Lost' stage in the Sales pipeline.")
    closedlost_stage_id = closedlost_stage["id"]
    closedlost_date_prop_name = f"hs_v2_date_entered_{closedlost_stage_id}"
    if closedlost_date_prop_name not in deal_props_by_name:
        raise HubSpotError(f"Expected deal property '{closedlost_date_prop_name}' does not exist.")
    print(f"  Closed Lost stage id                 = {closedlost_stage_id} (label: {closedlost_stage.get('label')})")
    print(f"  Closed Lost date-entered property    = {closedlost_date_prop_name} "
          f"(resolved per spec; not referenced by any funnel/stage formula below)")

    # --- SQL Lost Reason
    sql_lost_prop = _by_label(deal_props, "SQL Lost Reason")
    if not sql_lost_prop:
        raise HubSpotError("Could not resolve the 'SQL Lost Reason' deal property by exact label.")
    if sql_lost_prop["name"] != "sql_lost_reason":
        msg = (f"NOTE: spec's guessed internal name 'sql_lost_reason' does not exist; using "
               f"'{sql_lost_prop['name']}' (label 'SQL Lost Reason', exact match).")
        print(f"  {msg}")
        flags.append(msg)
    unreachable_values = _match_option_values(sql_lost_prop, ["Unreachable"], flags)
    if "Unreachable" not in unreachable_values:
        raise HubSpotError(f"'{sql_lost_prop['name']}' has no option 'Unreachable'.")
    print(f"  SQL Lost Reason                      = {sql_lost_prop['name']} "
          f"(Unreachable value = {unreachable_values['Unreachable']!r})")

    # --- hubspot_owner_assigneddate (standard property)
    owner_assigneddate_prop = deal_props_by_name.get("hubspot_owner_assigneddate")
    if not owner_assigneddate_prop:
        raise HubSpotError("Standard deal property 'hubspot_owner_assigneddate' does not exist on this portal.")
    print(f"  Owner assigned date                  = {owner_assigneddate_prop['name']}")

    # --- Proposal Sent / Signed Date Time
    proposal_sent_prop = _by_label_contains(deal_props, "proposal", "sent")
    proposal_signed_prop = _by_label_contains(deal_props, "proposal", "sign")
    if not proposal_sent_prop or not proposal_signed_prop:
        raise HubSpotError("Could not resolve 'Proposal Sent/Signed Date Time' deal properties.")
    if proposal_sent_prop["name"] != "proposal_sent_date":
        msg = (f"NOTE: spec's guessed name 'proposal_sent_date' does not exist; using "
               f"'{proposal_sent_prop['name']}' (label '{proposal_sent_prop.get('label')}').")
        print(f"  {msg}")
        flags.append(msg)
    if proposal_signed_prop["name"] != "proposal_signed_date":
        msg = (f"NOTE: spec's guessed name 'proposal_signed_date' does not exist; using "
               f"'{proposal_signed_prop['name']}' (label '{proposal_signed_prop.get('label')}').")
        print(f"  {msg}")
        flags.append(msg)
    print(f"  Proposal Sent Date Time              = {proposal_sent_prop['name']}")
    print(f"  Proposal Signed Date Time            = {proposal_signed_prop['name']}")

    print("=== Step 0 complete ===\n")

    return ReferenceData(
        source_prop=source_prop["name"], dd1_prop=dd1_prop["name"], dd2_prop=dd2_prop["name"],
        contact_type_prop=contact_type_prop["name"], contact_type_exclude_values=contact_type_exclude_values,
        brand_prop=brand_prop["name"], brand_required_value=brand_values[BRAND_REQUIRED_LABEL],
        brand_domain_prop=brand_domain_prop["name"], brand_domain_exclude_values=brand_domain_exclude_values,
        lead_source_prop=lead_source_prop["name"], lead_source_exclude_values=lead_source_exclude_values,
        sales_pipeline_id=sales_pipeline["id"], sales_pipeline_label=sales_pipeline.get("label", ""),
        sql_lost_prop=sql_lost_prop["name"], owner_assigneddate_prop=owner_assigneddate_prop["name"],
        closedlost_stage_id=closedlost_stage_id, closedlost_date_entered_prop=closedlost_date_prop_name,
        proposal_sent_prop=proposal_sent_prop["name"], proposal_signed_prop=proposal_signed_prop["name"],
        flags=flags,
    )


# =========================================================================
# Data fetch
# =========================================================================


def _resolve_fetch_cutoff(run_date: date) -> tuple[datetime, bool]:
    """Returns (cutoff, from_env). FETCH_CONTACTS_SINCE (.env, YYYY-MM-DD)
    overrides the default of Jan 1 of the report year."""
    since = os.environ.get("FETCH_CONTACTS_SINCE", "").strip()
    if since:
        try:
            return datetime.strptime(since, "%Y-%m-%d").replace(tzinfo=timezone.utc), True
        except ValueError:
            raise HubSpotError(f"FETCH_CONTACTS_SINCE={since!r} is not a valid YYYY-MM-DD date.")
    return datetime(run_date.year, 1, 1, tzinfo=timezone.utc), False


def _overall_filter_steps(ref: ReferenceData, cutoff_ms: str) -> list[tuple[str, dict, bool]]:
    """The 4 overall filters as individual named search-filter clauses, in
    the same order fetch_contacts() ANDs them together -- shared with
    debug_contact_fetch_filters() so the incremental diagnostic tests the
    exact same clauses the real fetch uses, not a re-typed approximation.

    Each entry is (label, clause, server_side). Brand was previously
    server_side=False, based on an earlier DEBUG_FETCH_FILTERS run that
    showed the filtered count collapsing (21,616 -> 37 for Brand, vs.
    ~13,617 expected). That diagnosis turned out to be a red herring: at
    the time, Brand had been mis-resolved to the wrong property
    (client_x_brand, ~0.06% populated -- see resolve_reference_data's
    Brand section for why) because hs_all_assigned_business_unit_ids's
    "options" array is always empty ("externalOptions": true) and the old
    resolution logic required a populated options match. So that CONTAINS_TOKEN
    filter was silently testing client_x_brand's own (near-empty)
    population the whole time, not a real Business Units filtering
    restriction on any credential. Now that Brand resolves correctly,
    Brand is server_side=True like the other 3 filters -- narrowing the
    search server-side is strictly better (fewer results paginated,
    less risk of hitting the Search API's 10,000-result cap).
    apply_overall_filters() still re-applies it client-side too, as the
    authoritative pass this whole function's docstring already promises."""
    return [
        ("createdate cutoff", {"propertyName": "createdate", "operator": "GTE", "value": cutoff_ms}, True),
        ("+ Brand (required)", {"propertyName": ref.brand_prop, "operator": "CONTAINS_TOKEN",
                                 "value": ref.brand_required_value}, True),
        ("+ Contact Type (exclude)", {"propertyName": ref.contact_type_prop, "operator": "NOT_IN",
                                       "values": list(ref.contact_type_exclude_values.values())}, True),
        ("+ Lead Source (exclude)", {"propertyName": ref.lead_source_prop, "operator": "NOT_IN",
                                      "values": list(ref.lead_source_exclude_values.values())}, True),
        ("+ Brand Domain (exclude)", {"propertyName": ref.brand_domain_prop, "operator": "NOT_IN",
                                       "values": list(ref.brand_domain_exclude_values.values())}, True),
    ]


def debug_contact_fetch_filters(client: HubSpotClient, ref: ReferenceData, run_date: date) -> None:
    """Diagnostic, opt-in via DEBUG_FETCH_FILTERS=1 in .env: adds the 4
    overall filters to the contact search ONE AT A TIME, printing
    HubSpot's own reported `total` after each addition. Pinpoints exactly
    which filter (if any) causes an unexpected drop for this specific
    credential/token, rather than guessing -- run this before the real
    fetch if the final contact count looks implausibly low."""
    cutoff, from_env = _resolve_fetch_cutoff(run_date)
    cutoff_ms = str(int(cutoff.timestamp() * 1000))
    steps = _overall_filter_steps(ref, cutoff_ms)
    print(f"=== DEBUG_FETCH_FILTERS: incremental filter totals (cutoff {cutoff.date().isoformat()}"
          f"{', from FETCH_CONTACTS_SINCE' if from_env else ''}) ===")
    accumulated: list[dict] = []
    for label, clause, _server_side in steps:
        accumulated.append(clause)
        body = {"filterGroups": [{"filters": list(accumulated)}], "properties": ["createdate"], "limit": 1}
        data = client.post("/crm/v3/objects/contacts/search", is_search=True, json=body)
        total = data.get("total")
        print(f"  {label:<28} total={total:,}" if isinstance(total, int) else f"  {label:<28} total={total!r}")
    print("=== end DEBUG_FETCH_FILTERS ===\n")


def fetch_contacts(client: HubSpotClient, ref: ReferenceData, run_date: date) -> list[dict]:
    """Fetches contacts created on/after a cutoff date, pre-filtered
    server-side by the 4 overall filters (still re-applied client-side in
    apply_overall_filters() as the authoritative pass -- this is purely a
    volume-reduction optimization, not a substitute).

    Cutoff date: `FETCH_CONTACTS_SINCE` in `.env` (format YYYY-MM-DD)
    overrides the default of Jan 1 of the report year -- e.g. set it to
    2026-04-01 to only fetch contacts created on/after that date. A later
    cutoff means a smaller, faster fetch, at the cost of the trade-off
    below getting worse (fewer prior contacts are considered at all).

    ACCEPTED TRADE-OFF (explicitly requested, not a spec default): the
    New Deals and Existing Deals funnels are defined around contacts
    created *before* the period being evaluated, so a contact created
    before the cutoff that still produced deal activity this report year
    will not appear in either funnel -- only in New Contacts (which
    requires contact.createdate this period anyway, so it's unaffected).

    Also note: HubSpot's Search API caps total results at 10,000
    regardless of pagination (see search_all()'s docstring) -- narrow the
    cutoff date if the printed fetched-count hits that cap.
    """
    props = [ref.source_prop, ref.dd1_prop, ref.dd2_prop, ref.contact_type_prop,
             ref.brand_prop, ref.brand_domain_prop, ref.lead_source_prop, "createdate"]

    cutoff, from_env = _resolve_fetch_cutoff(run_date)
    # HubSpot's Search API requires filter values as strings, even for a
    # numeric epoch-ms datetime filter -- sending a JSON integer here
    # previously raised a 400 ("problem with the request").
    cutoff_ms = str(int(cutoff.timestamp() * 1000))
    all_steps = _overall_filter_steps(ref, cutoff_ms)
    skipped = [label for label, _clause, server_side in all_steps if not server_side]
    print(f"  Fetching contacts created on/after {cutoff.date().isoformat()}"
          f"{' (FETCH_CONTACTS_SINCE)' if from_env else ' (report-year default)'}, "
          f"pre-filtered by the overall filters...")
    if skipped:
        print(f"  NOTE: not filtering server-side on: {', '.join(skipped)} -- confirmed this Service Key's "
              f"Search API results collapse when filtering on the Business Units (Brand) property "
              f"specifically (a per-credential restriction, not a bug in the property mapping). "
              f"apply_overall_filters() still applies it client-side as the authoritative check.")

    if os.environ.get("DEBUG_FETCH_FILTERS", "").strip().lower() in ("1", "true", "yes"):
        debug_contact_fetch_filters(client, ref, run_date)

    filters = [clause for _label, clause, server_side in all_steps if server_side]
    body = {"filterGroups": [{"filters": filters}], "properties": props}
    raw = client.search_all("contacts", body, label="contacts")
    contacts = []
    for r in raw:
        p = r.get("properties", {})
        createdate = parse_hs_datetime(p.get("createdate"))
        if createdate is None:
            continue  # a contact with no createdate can't be bucketed by any funnel; excluded, not silently kept
        contacts.append({
            "id": r["id"],
            "createdate": createdate,
            "source": p.get(ref.source_prop) or None,
            "dd1": p.get(ref.dd1_prop) or None,
            "dd2": p.get(ref.dd2_prop) or None,
            "contact_type": p.get(ref.contact_type_prop) or None,
            "brand": p.get(ref.brand_prop) or None,
            "brand_domain": p.get(ref.brand_domain_prop) or None,
            "lead_source": p.get(ref.lead_source_prop) or None,
            "deals": [],
        })
    return contacts


def attach_deals(client: HubSpotClient, ref: ReferenceData, contacts: list[dict]) -> None:
    contact_ids = [c["id"] for c in contacts]
    print("  Fetching contact-to-deal associations...")
    assoc_map = client.batch_read_associations("contacts", "deals", contact_ids, label="contact-deal associations")
    all_deal_ids = sorted({d for ids in assoc_map.values() for d in ids})
    deal_props = ["pipeline", "createdate", ref.sql_lost_prop, ref.owner_assigneddate_prop,
                  ref.proposal_sent_prop, ref.proposal_signed_prop]
    print(f"  Fetching properties for {len(all_deal_ids):,} distinct deals...")
    deals_raw = client.batch_read("deals", all_deal_ids, deal_props, label="deals")
    deals_by_id = {}
    for d in deals_raw:
        p = d.get("properties", {})
        createdate = parse_hs_datetime(p.get("createdate"))
        if createdate is None:
            continue
        deals_by_id[d["id"]] = {
            "id": d["id"],
            "createdate": createdate,
            "pipeline": p.get("pipeline") or None,
            "sql_lost": p.get(ref.sql_lost_prop) or None,
            "owner_assigneddate": parse_hs_datetime(p.get(ref.owner_assigneddate_prop)),
            "proposal_sent": parse_hs_datetime(p.get(ref.proposal_sent_prop)),
            "proposal_signed": parse_hs_datetime(p.get(ref.proposal_signed_prop)),
        }
    for c in contacts:
        c["deals"] = [deals_by_id[d_id] for d_id in assoc_map.get(c["id"], []) if d_id in deals_by_id]


# =========================================================================
# Overall filters (applied sequentially; each keeps the audit trail spec asks for)
# =========================================================================


@dataclass
class FilterStep:
    description: str
    count_before: int
    count_after: int
    excluded_samples: list
    blank_pass_candidates: list  # narrowed to final-output survivors after all 4 filters run


def apply_overall_filters(contacts: list[dict], ref: ReferenceData) -> tuple[list[dict], list[FilterStep]]:
    steps = []
    current = contacts
    print("=== Overall filter audit trail ===")
    print(f"  Start: {len(current)} contacts")

    def run_filter(desc, keep_fn, blank_fn):
        nonlocal current
        before = len(current)
        excluded = [c for c in current if not keep_fn(c)][:3]
        kept = [c for c in current if keep_fn(c)]
        # Every blank-on-this-filter survivor is kept as a *candidate*; a
        # blank passing this filter can still be excluded by one of the
        # other 3, so we can't fix the sample of 2-3 until the full
        # pipeline has run (see the narrowing pass below run_filter calls).
        blank_pass_candidates = [c for c in kept if blank_fn(c)]
        current = kept
        steps.append(FilterStep(desc, before, len(current), excluded, blank_pass_candidates))
        print(f"  After [{desc}]: {len(current)} contacts (was {before})")

    excl1 = set(ref.contact_type_exclude_values.values())
    run_filter(
        "Contact Type excludes B2B Partnership Development / B2B Institutional Relations (blanks pass)",
        keep_fn=lambda c: c["contact_type"] is None or c["contact_type"] not in excl1,
        blank_fn=lambda c: c["contact_type"] is None,
    )

    excl2 = set(ref.lead_source_exclude_values.values())
    run_filter(
        "Lead Source excludes Bundle Offer/Other/Instantly/Private/Walk-In/Email/Partner Referral/"
        "Phone Calls/Events/Client Referral (blanks pass)",
        keep_fn=lambda c: c["lead_source"] is None or c["lead_source"] not in excl2,
        blank_fn=lambda c: c["lead_source"] is None,
    )

    run_filter(
        f"Brand includes '{ref.brand_required_value}' (blanks FAIL this filter)",
        keep_fn=lambda c: _has_brand_value(c["brand"], ref.brand_required_value),
        blank_fn=lambda c: False,  # no blank-pass case for this filter by design
    )

    excl4 = set(ref.brand_domain_exclude_values.values())
    run_filter(
        "Brand Domain excludes 'BePortugal' (blanks pass)",
        keep_fn=lambda c: c["brand_domain"] is None or c["brand_domain"] not in excl4,
        blank_fn=lambda c: c["brand_domain"] is None,
    )

    final_ids = {c["id"] for c in current}
    for step in steps:
        step.blank_pass_candidates = [c for c in step.blank_pass_candidates if c["id"] in final_ids][:3]

    print(f"=== Overall filters complete: {len(current)} contacts remain ===\n")
    return current, steps


# =========================================================================
# Row hierarchy (sparse, data-driven)
# =========================================================================


NOT_SET = "(not set)"


def build_rows(filtered_contacts: list[dict]) -> list[dict]:
    groups: dict = {}
    for c in filtered_contacts:
        source = c["source"] or NOT_SET
        dd1 = c["dd1"] or NOT_SET
        dd2 = c["dd2"] or NOT_SET
        groups.setdefault(source, {}).setdefault(dd1, {}).setdefault(dd2, []).append(c)

    rows = []
    for source in sorted(groups):
        l1_contacts = [c for dd1v in groups[source].values() for dd2v in dd1v.values() for c in dd2v]
        rows.append({"level": 1, "source": source, "dd1": None, "dd2": None, "contacts": l1_contacts})
        for dd1 in sorted(groups[source]):
            l2_contacts = [c for dd2v in groups[source][dd1].values() for c in dd2v]
            rows.append({"level": 2, "source": source, "dd1": dd1, "dd2": None, "contacts": l2_contacts})
            for dd2 in sorted(groups[source][dd1]):
                l3_contacts = groups[source][dd1][dd2]
                rows.append({"level": 3, "source": source, "dd1": dd1, "dd2": dd2, "contacts": l3_contacts})
    return rows


# =========================================================================
# Funnel/stage qualification -- bucketing is done at Month granularity
# (matching the spec's literal "this month"/"before this month" wording);
# Week/Day are a pure sub-split of the same already-qualified population by
# the stage's own bucketing date, which is what guarantees Month == sum of
# its Weeks == sum of their Days (QA check #1/#2) by construction rather
# than by coincidence.
# =========================================================================


def _stage_nondate_ok(ref: ReferenceData, funnel_name: str, stage_idx: int, deal: dict) -> bool:
    stage = STAGE_LABELS[stage_idx]
    if stage == "New Deals":
        return True  # no pipeline restriction at this stage, in either funnel
    if deal["pipeline"] != ref.sales_pipeline_id:
        return False
    if stage == "Qualified Deals":
        return True
    if stage == "Opportunities":
        return deal["sql_lost"] != "Unreachable"  # None (blank) also passes
    if stage == "Proposals Sent":
        return deal["proposal_sent"] is not None
    if stage == "Proposal Signed":
        return deal["proposal_signed"] is not None
    return False


def _existing_deals_milestone(stage_idx: int, deal: dict) -> datetime | None:
    stage = STAGE_LABELS[stage_idx]
    if stage in ("Qualified Deals", "Opportunities"):
        return deal["owner_assigneddate"]
    if stage == "Proposals Sent":
        return deal["proposal_sent"]
    if stage == "Proposal Signed":
        return deal["proposal_signed"]
    return None


def stage_bucket_day(ref: ReferenceData, funnel_name: str, stage_idx: int, contact: dict, deal: dict) -> date | None:
    """Returns the Day this (contact, deal) pair contributes to for the
    given funnel/stage's deal-counted metric, or None if it doesn't qualify."""
    if funnel_name == "New Contacts":
        deal_day = deal["createdate"].date()
        m_start, m_end = month_start(deal_day), month_end_exclusive(deal_day)
        if not (m_start <= contact["createdate"].date() < m_end):
            return None
        if not _stage_nondate_ok(ref, funnel_name, stage_idx, deal):
            return None
        return deal_day

    if funnel_name == "New Deals":
        deal_day = deal["createdate"].date()
        m_start = month_start(deal_day)
        if not (contact["createdate"].date() < m_start):
            return None
        if not _stage_nondate_ok(ref, funnel_name, stage_idx, deal):
            return None
        return deal_day

    # Existing Deals
    milestone = _existing_deals_milestone(stage_idx, deal)
    if milestone is None:
        return None
    milestone_day = milestone.date()
    m_start = month_start(milestone_day)
    if not (contact["createdate"].date() < m_start and deal["createdate"].date() < m_start):
        return None
    if not _stage_nondate_ok(ref, funnel_name, stage_idx, deal):
        return None
    return milestone_day


def new_contacts_stage_bucket(contact: dict) -> date | None:
    """Contact-dimension bucket for Funnel: New Contacts / Stage: New Contacts."""
    c_day = contact["createdate"].date()
    m_start, m_end = month_start(c_day), month_end_exclusive(c_day)
    for deal in contact["deals"]:
        if m_start <= deal["createdate"].date() < m_end:
            return c_day
    return None


def compute_leaf_counts(rows: list[dict], ref: ReferenceData) -> dict:
    """For every Level-3 row and every applicable (funnel, stage), a
    {date: count} dict of literal (not formula) values. This is the single
    ground truth every rollup -- row-wise and time-wise -- is built from."""
    leaf_counts: dict = {}
    leaf_rows = [row for row in rows if row["level"] == 3]
    for processed, row in enumerate(leaf_rows, start=1):
        key = (row["source"], row["dd1"], row["dd2"])
        leaf_counts[key] = {}
        _print_progress("leaf rows computed", processed, total=len(leaf_rows))
        for funnel in FUNNELS:
            fname = funnel["name"]
            for stage_idx, applicable in enumerate(funnel["applicable"]):
                if not applicable:
                    continue
                counts: dict = {}
                if fname == "New Contacts" and stage_idx == 0:
                    for c in row["contacts"]:
                        d = new_contacts_stage_bucket(c)
                        if d is not None:
                            counts[d] = counts.get(d, 0) + 1
                else:
                    for c in row["contacts"]:
                        for deal in c["deals"]:
                            d = stage_bucket_day(ref, fname, stage_idx, c, deal)
                            if d is not None:
                                counts[d] = counts.get(d, 0) + 1
                leaf_counts[key][(fname, stage_idx)] = counts
    if leaf_rows:
        _print_progress_done()
    return leaf_counts


# =========================================================================
# Time-column layout (Month -> Week -> Day, year-to-date, identical
# shape for every one of the 18 stage-groups)
# =========================================================================


@dataclass
class DayCol:
    col: int
    day: date
    has_data: bool


@dataclass
class WeekCol:
    label: str
    col: int
    days: list
    is_complete: bool


@dataclass
class MonthCol:
    idx: int
    name: str
    col: int
    weeks: list
    is_complete: bool


def build_time_layout(run_date: date) -> tuple[list, int]:
    """Year-to-date only: January through the current month (run_date's
    month), not the full calendar year -- a month after run_date simply
    isn't rendered as a column at all, rather than existing with entirely
    blank cells."""
    year = run_date.year
    months: list = []
    col = 1  # 1-based, relative to the start of a stage-group's column block
    for month_idx in range(1, run_date.month + 1):
        first = date(year, month_idx, 1)
        last = month_end_exclusive(first) - timedelta(days=1)
        is_month_complete = last < month_start(run_date) if run_date.year == year else True
        weeks: list = []
        raw_weeks: list = []
        cur_week_start, cur_days = None, []
        d_ord = first.toordinal()
        while d_ord <= last.toordinal():
            day = date.fromordinal(d_ord)
            wk = week_start(day)
            if wk != cur_week_start:
                if cur_days:
                    raw_weeks.append(cur_days)
                cur_week_start, cur_days = wk, []
            cur_days.append(day)
            d_ord += 1
        if cur_days:
            raw_weeks.append(cur_days)

        for wi, days in enumerate(raw_weeks):
            day_cols = []
            for day in days:
                has_data = day <= run_date
                day_cols.append(DayCol(col=col, day=day, has_data=has_data))
                col += 1
            week_col = col
            week_is_complete = days[-1] < week_start(run_date)
            weeks.append(WeekCol(label=f"W{wi + 1}", col=week_col, days=day_cols, is_complete=week_is_complete))
            col += 1

        month_col = col
        months.append(MonthCol(idx=month_idx, name=MONTH_NAMES[month_idx - 1], col=month_col,
                                weeks=weeks, is_complete=is_month_complete))
        col += 1

    return months, col - 1


LABEL_COL = 1  # column A holds the row label


# =========================================================================
# Workbook construction
# =========================================================================


def build_workbook(rows: list[dict], leaf_counts: dict, ref: ReferenceData,
                    run_date: date) -> tuple[Workbook, dict, dict]:
    """One sheet per funnel (no combined sheet, no Filters & Definitions
    sheet -- see the "3 per-funnel tabs" follow-up). Returns
    (workbook, {sheet_name: {cell_ref: cached_value}}, {sheet_name: column_count})."""
    months, group_width = build_time_layout(run_date)
    wb = Workbook()
    print("  Pre-computing row/time rollups (for cell formulas' cached values)...")
    computed = recompute_all_cells(rows, leaf_counts, ref, run_date)

    sheet_cell_values: dict = {}
    sheet_col_counts: dict = {}
    for i, funnel in enumerate(FUNNELS):
        print(f"  Building sheet '{funnel['name']}' ({i + 1}/{len(FUNNELS)})...")
        ws = wb.active if i == 0 else wb.create_sheet(funnel["name"])
        ws.title = funnel["name"]
        formula_cells: dict = {}
        col_count = _build_single_funnel_sheet(ws, funnel, rows, leaf_counts, computed, run_date,
                                                months, group_width, formula_cells)
        sheet_cell_values[funnel["name"]] = formula_cells
        sheet_col_counts[funnel["name"]] = col_count

    return wb, sheet_cell_values, sheet_col_counts


def _assign_row_numbers(rows: list[dict], header_rows: int) -> tuple[list[dict], dict]:
    """Attach a sheet row number to every row, and collect each Level-1/2
    row's immediate-children row numbers (for the row-rollup formulas)."""
    numbered = []
    child_rows: dict = {}  # index into `numbered` -> list of child row numbers
    r = header_rows + 1
    for row in rows:
        row = dict(row)
        row["row_num"] = r
        numbered.append(row)
        r += 1
    # second pass: find immediate children (next rows at level+1 until a
    # row at <= this row's level appears)
    for i, row in enumerate(numbered):
        if row["level"] == 3:
            continue
        children = []
        for j in range(i + 1, len(numbered)):
            other = numbered[j]
            if other["level"] <= row["level"]:
                break
            if other["level"] == row["level"] + 1:
                children.append(other["row_num"])
        child_rows[row["row_num"]] = children
    return numbered, child_rows


def _build_single_funnel_sheet(ws, funnel, rows, leaf_counts, computed, run_date,
                                months, group_width, formula_cells) -> int:
    """Builds one funnel's own sheet: Stage -> Month -> Week -> Day columns
    (no funnel band -- the tab itself is the funnel), only that funnel's
    applicable stages (N/A ones are omitted entirely, not placeholder-filled,
    per the "3 per-funnel tabs" follow-up)."""
    fname = funnel["name"]
    # summaryBelow=False: Level-1 rows are written ABOVE their Level-2/3
    # children in this sheet (not below, as Excel's default subtotal-style
    # grouping assumes) -- summaryBelow=True here would give Excel the
    # wrong anchor row for each group's +/- control and is what "click +
    # expands downward" actually requires for a summary-row-on-top layout.
    # summaryRight=True is correct as-is: Week/Month totals sit to the
    # right of the columns they sum.
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)

    HEADER_ROWS = 2  # 1: stage band, 2: time label
    numbered_rows, child_rows = _assign_row_numbers(rows, HEADER_ROWS)

    # --- corner title
    ws.merge_cells(start_row=1, start_column=LABEL_COL, end_row=HEADER_ROWS, end_column=LABEL_COL)
    corner = ws.cell(row=1, column=LABEL_COL, value=f"Original Traffic Source ({run_date.year})")
    corner.fill = HEADER_FILL
    corner.font = CORNER_TITLE_FONT
    corner.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    corner.border = THIN_BORDER
    ws.column_dimensions[get_column_letter(LABEL_COL)].width = 30

    # --- row labels + row outline levels + row-side rollup formulas (Day cells only)
    #
    # Only Level 1 is visible on open -- Level 2 AND Level 3 both start
    # hidden. (A prior version wrongly left Level 2 visible by default,
    # which made opening the sheet look like two levels were already
    # expanded, and a single click on Level 1 jumped straight to Level 3.)
    for row in numbered_rows:
        r = row["row_num"]
        if row["level"] == 1:
            label = row["source"]
        elif row["level"] == 2:
            label = f"    {row['dd1']}"
        else:
            label = f"        {row['dd2']}"
        cell = ws.cell(row=r, column=LABEL_COL, value=label)
        cell.font = BOLD_BODY_FONT if row["level"] == 1 else BODY_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(indent=row["level"] - 1)
        ws.row_dimensions[r].outline_level = row["level"] - 1
        ws.row_dimensions[r].hidden = row["level"] != 1
        if row["level"] in (1, 2):
            ws.row_dimensions[r].collapsed = True  # its children start collapsed

    # --- column layout + data cells, one applicable stage-group at a time
    ramp = hsl_ramp(funnel["base_color"], n=6)
    applicable_stages = [(i, s) for i, s in enumerate(STAGE_LABELS) if funnel["applicable"][i]]
    col = LABEL_COL + 1
    total_data_cols = 0
    for done, (stage_idx, stage_label) in enumerate(applicable_stages, start=1):
        _print_progress(f"stage columns written ({stage_label})", done, total=len(applicable_stages))
        stage_start_col = col
        fill_hex = ramp[stage_idx]
        stage_fill = PatternFill(fill_type="solid", fgColor=fill_hex)
        # Qualified Deals' ramp step computes as light enough (>~55%) for
        # readable_text_color() to pick Night Blue text automatically, but
        # against the Electric Blue hue specifically that's still hard to
        # read -- forced to white here per request.
        text_color = "FFFFFF" if stage_label == "Qualified Deals" else readable_text_color(fill_hex)
        stage_font = Font(name=FONT_BODY, bold=True, color=text_color)

        for month in months:
            for week in month.weeks:
                for day_col in week.days:
                    c = stage_start_col + day_col.col - 1
                    _write_time_header(ws, c, MONO_FONT, day_col.day.isoformat())
                    ws.column_dimensions[get_column_letter(c)].outline_level = 2
                    ws.column_dimensions[get_column_letter(c)].hidden = True
                    if day_col.has_data:
                        _write_day_value(ws, numbered_rows, child_rows, leaf_counts, computed,
                                          fname, stage_idx, c, day_col.day, formula_cells)
                wc = stage_start_col + week.col - 1
                _write_time_header(ws, wc, MONO_FONT, week.label)
                ws.column_dimensions[get_column_letter(wc)].outline_level = 1
                ws.column_dimensions[get_column_letter(wc)].hidden = True
                ws.column_dimensions[get_column_letter(wc)].collapsed = True
                first_day_c = stage_start_col + week.days[0].col - 1
                last_day_c = stage_start_col + week.days[-1].col - 1
                _write_week_or_month_total(ws, numbered_rows, computed, fname, stage_idx,
                                            first_day_c, last_day_c, wc, formula_cells,
                                            bucket_key=("week", (month.idx, week.label)))
            mc = stage_start_col + month.col - 1
            _write_time_header(ws, mc, BOLD_BODY_FONT, month.name)
            ws.column_dimensions[get_column_letter(mc)].outline_level = 0
            ws.column_dimensions[get_column_letter(mc)].collapsed = True
            week_total_cols = [stage_start_col + wk.col - 1 for wk in month.weeks]
            _write_week_or_month_total(ws, numbered_rows, computed, fname, stage_idx,
                                        None, None, mc, formula_cells,
                                        explicit_cols=week_total_cols,
                                        bucket_key=("month", month.idx))

        stage_end_col = stage_start_col + group_width - 1
        ws.merge_cells(start_row=1, start_column=stage_start_col, end_row=1, end_column=stage_end_col)
        ws.cell(row=1, column=stage_start_col, value=stage_label)
        for c in range(stage_start_col, stage_end_col + 1):
            cell1 = ws.cell(row=1, column=c)
            cell1.fill = stage_fill
            cell1.font = stage_font
            cell1.border = THIN_BORDER
            cell1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        col = stage_end_col + 1
        total_data_cols += group_width

    if applicable_stages:
        _print_progress_done()

    last_col = col - 1
    for c in range(LABEL_COL + 1, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 11

    ws.freeze_panes = ws.cell(row=HEADER_ROWS + 1, column=LABEL_COL + 1).coordinate
    return total_data_cols


def _write_time_header(ws, col, font, value):
    cell = ws.cell(row=2, column=col, value=value)  # row 2: the time-label header row
    cell.font = font
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center")


def _row_key(row):
    return (row["level"], row["source"], row["dd1"], row["dd2"])


def _write_day_value(ws, numbered_rows, child_rows, leaf_counts, computed, fname, stage_idx, col, day, formula_cells):
    letter = get_column_letter(col)
    for row in numbered_rows:
        r = row["row_num"]
        if row["level"] == 3:
            key = (row["source"], row["dd1"], row["dd2"])
            value = leaf_counts.get(key, {}).get((fname, stage_idx), {}).get(day, 0)
            cell = ws.cell(row=r, column=col, value=value)
        else:
            children = child_rows.get(r, [])
            cached_value = computed.get(_row_key(row), {}).get((fname, stage_idx), {}).get("day", {}).get(day, 0)
            if not children:
                cell = ws.cell(row=r, column=col, value=0)
                formula_cells[f"{letter}{r}"] = 0
                cell.font = BODY_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center")
                continue
            refs = ",".join(f"{letter}{cr}" for cr in children)
            cell = ws.cell(row=r, column=col, value=f"=SUM({refs})")
            formula_cells[f"{letter}{r}"] = cached_value
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")


def _write_week_or_month_total(ws, numbered_rows, computed, fname, stage_idx, first_col, last_col, total_col,
                                formula_cells, explicit_cols=None, bucket_key=None):
    letter = get_column_letter(total_col)
    kind, bucket = bucket_key
    for row in numbered_rows:
        r = row["row_num"]
        if explicit_cols is not None:
            refs = ",".join(f"{get_column_letter(c)}{r}" for c in explicit_cols)
        else:
            refs = f"{get_column_letter(first_col)}{r}:{get_column_letter(last_col)}{r}"
        cell = ws.cell(row=r, column=total_col, value=f"=SUM({refs})")
        cached_value = computed.get(_row_key(row), {}).get((fname, stage_idx), {}).get(kind, {}).get(bucket, 0)
        formula_cells[f"{letter}{r}"] = cached_value
        cell.font = BOLD_BODY_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")


# =========================================================================
# Independent (pure-Python) recomputation of every rollup, used both for
# the recalculation cached-value fallback and for QA checks 1/2.
# =========================================================================


def recompute_all_cells(rows, leaf_counts, ref, run_date) -> dict:
    """Returns {(source, dd1, dd2, level): {(funnel, stage): {"day": {date: v}, "week": {...}, "month": {...}}}}
    purely in Python, used to cross-check the workbook's own formulas."""
    months, _width = build_time_layout(run_date)
    result: dict = {}

    def leaf_key(row):
        return (row["source"], row["dd1"], row["dd2"])

    # Level 3: literal leaf counts, rolled up to week/month in Python.
    for row in rows:
        if row["level"] != 3:
            continue
        key = leaf_key(row)
        result[(row["level"],) + key] = {}
        for funnel in FUNNELS:
            for stage_idx, applicable in enumerate(funnel["applicable"]):
                if not applicable:
                    continue
                day_counts = leaf_counts.get(key, {}).get((funnel["name"], stage_idx), {})
                week_counts, month_counts = {}, {}
                for month in months:
                    m_total = 0
                    for week in month.weeks:
                        w_total = sum(day_counts.get(d.day, 0) for d in week.days if d.has_data)
                        week_counts[(month.idx, week.label)] = w_total
                        m_total += w_total
                    month_counts[month.idx] = m_total
                result[(row["level"],) + key][(funnel["name"], stage_idx)] = {
                    "day": day_counts, "week": week_counts, "month": month_counts,
                }

    # Level 1/2: sum of immediate children's same-bucket values.
    for level in (2, 1):
        for row in rows:
            if row["level"] != level:
                continue
            key_prefix = (row["source"], row["dd1"]) if level == 2 else (row["source"],)
            children = [rr for rr in rows if rr["level"] == level + 1
                        and rr["source"] == row["source"]
                        and (level == 1 or rr["dd1"] == row["dd1"])]
            key = (row["source"], row["dd1"], row["dd2"])
            result[(level,) + key] = {}
            for funnel in FUNNELS:
                for stage_idx, applicable in enumerate(funnel["applicable"]):
                    if not applicable:
                        continue
                    day_counts, week_counts, month_counts = {}, {}, {}
                    for child in children:
                        child_key = (level + 1, child["source"], child["dd1"], child["dd2"])
                        child_data = result.get(child_key, {}).get((funnel["name"], stage_idx))
                        if not child_data:
                            continue
                        for d, v in child_data["day"].items():
                            day_counts[d] = day_counts.get(d, 0) + v
                        for wk, v in child_data["week"].items():
                            week_counts[wk] = week_counts.get(wk, 0) + v
                        for m, v in child_data["month"].items():
                            month_counts[m] = month_counts.get(m, 0) + v
                    result[(level,) + key][(funnel["name"], stage_idx)] = {
                        "day": day_counts, "week": week_counts, "month": month_counts,
                    }
    return result


# =========================================================================
# Funnel/stage descriptions -- printed by QA check 4, not written to any
# sheet (the "Filters & Definitions" tab was dropped; see README).
# =========================================================================

FUNNEL_STAGE_DESCRIPTIONS = {
    ("New Contacts", "New Contacts"): "CONTACT.createdate this period AND DEAL.createdate this period",
    ("New Contacts", "New Deals"): "same base population, no pipeline restriction",
    ("New Contacts", "Qualified Deals"): "same base population AND pipeline = Sales",
    ("New Contacts", "Opportunities"): "same base population AND pipeline = Sales AND SQL Lost Reason != Unreachable (blank passes)",
    ("New Contacts", "Proposals Sent"): "same base population AND pipeline = Sales AND Proposal Sent Date Time is known",
    ("New Contacts", "Proposal Signed"): "same base population AND pipeline = Sales AND Proposal Signed Date Time is known",
    ("New Deals", "New Contacts"): "N/A",
    ("New Deals", "New Deals"): "CONTACT.createdate before this period AND DEAL.createdate this period, any pipeline",
    ("New Deals", "Qualified Deals"): "same base population AND pipeline = Sales",
    ("New Deals", "Opportunities"): "same base population AND pipeline = Sales AND SQL Lost Reason != Unreachable (blank passes)",
    ("New Deals", "Proposals Sent"): "same base population AND pipeline = Sales AND Proposal Sent Date Time is known",
    ("New Deals", "Proposal Signed"): "same base population AND pipeline = Sales AND Proposal Signed Date Time is known",
    ("Existing Deals", "New Contacts"): "N/A",
    ("Existing Deals", "New Deals"): "N/A",
    ("Existing Deals", "Qualified Deals"): "CONTACT.createdate & DEAL.createdate before this period AND pipeline = Sales AND Owner Assigned Date falls in this period",
    ("Existing Deals", "Opportunities"): "same base AND SQL Lost Reason != Unreachable (blank passes) AND Owner Assigned Date falls in this period",
    ("Existing Deals", "Proposals Sent"): "same base AND pipeline = Sales AND Proposal Sent Date Time falls in this period",
    ("Existing Deals", "Proposal Signed"): "same base AND pipeline = Sales AND Proposal Signed Date Time falls in this period",
}


# =========================================================================
# Recalculation (LibreOffice round-trip, falling back to cached-value injection)
# =========================================================================


def _try_libreoffice_recalculation(xlsx_path: str) -> bool:
    out_dir = tempfile.mkdtemp(prefix="gcs_funnel_recalc_")
    try:
        result = subprocess.run(
            ["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out_dir, xlsx_path],
            capture_output=True, text=True, timeout=300,
        )
        recalculated = os.path.join(out_dir, os.path.basename(xlsx_path))
        if result.returncode != 0 or not os.path.exists(recalculated):
            print(f"  LibreOffice unavailable/failed ({result.stderr.strip() or result.stdout.strip()}), "
                  f"falling back to direct cached-value injection.")
            return False
        os.replace(recalculated, xlsx_path)
        return True
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"  LibreOffice unavailable ({exc}), falling back to direct cached-value injection.")
        return False
    finally:
        if os.path.isdir(out_dir):
            for f in os.listdir(out_dir):
                os.remove(os.path.join(out_dir, f))
            os.rmdir(out_dir)


def _inject_cached_formula_values(xlsx_path: str, sheet_cell_values: dict) -> None:
    import xml.etree.ElementTree as ET

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    doc_rels_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", ns)

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        workbook_xml = ET.fromstring(zin.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.get("Id"): rel.get("Target")
                          for rel in rels_xml.findall(f"{{{pkg_rels_ns}}}Relationship")}

        sheet_name_to_xml = {}
        for sheet in workbook_xml.find(f"{{{ns}}}sheets"):
            name = sheet.get("name")
            rid = sheet.get(f"{{{doc_rels_ns}}}id")
            target = rid_to_target.get(rid)
            if name and target:
                sheet_name_to_xml[name] = target.lstrip("/") if target.startswith("/") else "xl/" + target

        updates = {}
        for sheet_name, cell_values in sheet_cell_values.items():
            xml_path = sheet_name_to_xml.get(sheet_name)
            if not xml_path or not cell_values:
                continue
            root = ET.fromstring(zin.read(xml_path))
            for c in root.iter(f"{{{ns}}}c"):
                ref = c.get("r")
                if ref in cell_values:
                    v_elem = c.find(f"{{{ns}}}v")
                    if v_elem is None:
                        v_elem = ET.SubElement(c, f"{{{ns}}}v")
                    v_elem.text = str(cell_values[ref])
            updates[xml_path] = ET.tostring(root, xml_declaration=True, encoding="UTF-8")

        all_names = zin.namelist()
        buffers = {name: zin.read(name) for name in all_names}

    for path, new_bytes in updates.items():
        buffers[path] = new_bytes

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in buffers.items():
            zout.writestr(name, data)


def recalculate_workbook(xlsx_path: str, sheet_cell_values: dict) -> None:
    if not _try_libreoffice_recalculation(xlsx_path):
        _inject_cached_formula_values(xlsx_path, sheet_cell_values)


# =========================================================================
# QA
# =========================================================================


def run_qa(rows, leaf_counts, ref, run_date, filtered_contacts, all_contacts, filter_steps, sheet_col_counts) -> bool:
    print("=== QA ===")
    computed = recompute_all_cells(rows, leaf_counts, ref, run_date)
    all_pass = True

    # --- QA 1: full click-path simulation for the richest Level-1 row
    l1_rows = [r for r in rows if r["level"] == 1]
    richest = max(l1_rows, key=lambda r: len(r["contacts"]), default=None)
    check1_pass = True
    check1_detail = []
    if richest:
        key = (1, richest["source"], None, None)
        stage_data = computed.get(key, {}).get(("New Contacts", 2))  # Qualified Deals under New Contacts funnel
        months, _w = build_time_layout(run_date)
        finished_month = next((m for m in months if m.is_complete and stage_data), None)
        in_progress_month = next((m for m in months if m.idx == run_date.month), None)
        for month in filter(None, [finished_month, in_progress_month]):
            if not stage_data:
                break
            month_total = stage_data["month"].get(month.idx, 0)
            week_sum = sum(stage_data["week"].get((month.idx, w.label), 0) for w in month.weeks)
            if month_total != week_sum:
                check1_pass = False
                check1_detail.append(f"{richest['source']} / New Contacts-Qualified Deals / {month.name}: "
                                      f"month={month_total} != sum(weeks)={week_sum}")
            for w in month.weeks:
                week_total = stage_data["week"].get((month.idx, w.label), 0)
                day_sum = sum(stage_data["day"].get(d.day, 0) for d in w.days if d.has_data)
                if week_total != day_sum:
                    check1_pass = False
                    check1_detail.append(f"{richest['source']} / {month.name} {w.label}: "
                                          f"week={week_total} != sum(days)={day_sum}")
    all_pass &= check1_pass
    print(f"  [1] Full click-path simulation: {'PASS' if check1_pass else 'FAIL'}")
    for d in check1_detail:
        print(f"      - {d}")

    # --- QA 2: row rollup integrity, every row, every funnel/stage/bucket
    check2_pass = True
    check2_detail = []
    for row in rows:
        if row["level"] == 3:
            continue
        key = (row["level"], row["source"], row["dd1"], row["dd2"])
        children = [rr for rr in rows if rr["level"] == row["level"] + 1
                    and rr["source"] == row["source"]
                    and (row["level"] == 1 or rr["dd1"] == row["dd1"])]
        for funnel in FUNNELS:
            for stage_idx, applicable in enumerate(funnel["applicable"]):
                if not applicable:
                    continue
                parent_months = computed.get(key, {}).get((funnel["name"], stage_idx), {}).get("month", {})
                child_month_sum = {}
                for child in children:
                    ckey = (child["level"], child["source"], child["dd1"], child["dd2"])
                    cdata = computed.get(ckey, {}).get((funnel["name"], stage_idx), {}).get("month", {})
                    for m, v in cdata.items():
                        child_month_sum[m] = child_month_sum.get(m, 0) + v
                for m in set(parent_months) | set(child_month_sum):
                    if parent_months.get(m, 0) != child_month_sum.get(m, 0):
                        check2_pass = False
                        check2_detail.append(
                            f"level {row['level']} row {row['source']}/{row['dd1']}/{row['dd2']}, "
                            f"{funnel['name']}/{STAGE_LABELS[stage_idx]}, month {m}: "
                            f"row={parent_months.get(m, 0)} != sum(children)={child_month_sum.get(m, 0)}")
    all_pass &= check2_pass
    print(f"  [2] Row rollup integrity (all rows): {'PASS' if check2_pass else 'FAIL'}")
    for d in check2_detail[:10]:
        print(f"      - {d}")
    if len(check2_detail) > 10:
        print(f"      ... and {len(check2_detail) - 10} more")

    # --- QA 3: filter correctness spot-check
    check3_pass = True
    check3_detail = []
    filtered_ids = {c["id"] for c in filtered_contacts}
    for step in filter_steps:
        for c in step.excluded_samples:
            if c["id"] in filtered_ids:
                check3_pass = False
                check3_detail.append(f"[{step.description}] excluded sample contact {c['id']} still present in output")
        for c in step.blank_pass_candidates:
            if c["id"] not in filtered_ids:
                check3_pass = False
                check3_detail.append(f"[{step.description}] blank-pass sample contact {c['id']} missing from output")
    all_pass &= check3_pass
    print(f"  [3] Filter correctness spot-check: {'PASS' if check3_pass else 'FAIL'}")
    for d in check3_detail:
        print(f"      - {d}")

    # --- QA 4: funnel/stage definition correctness (printed totals)
    print("  [4] Funnel/Stage definition totals:")
    summary_rows = []
    for funnel in FUNNELS:
        for stage_idx, stage_label in enumerate(STAGE_LABELS):
            applicable = funnel["applicable"][stage_idx]
            if not applicable:
                print(f"      {funnel['name']:<16} {stage_label:<18} N/A")
                summary_rows.append((funnel["name"], stage_label, "N/A"))
                continue
            total = 0
            for row in rows:
                if row["level"] != 1:
                    continue
                key = (1, row["source"], None, None)
                month_data = computed.get(key, {}).get((funnel["name"], stage_idx), {}).get("month", {})
                total += sum(month_data.values())
            desc = FUNNEL_STAGE_DESCRIPTIONS[(funnel["name"], stage_label)]
            print(f"      {funnel['name']:<16} {stage_label:<18} total={total:<8} filter: {desc}")
            summary_rows.append((funnel["name"], stage_label, total))
    check4_pass = True  # this check is a printed report, not a pass/fail assertion

    # --- QA 5: to-date vs completed logic
    months, _w = build_time_layout(run_date)
    in_progress_month = next((m for m in months if m.idx == run_date.month), None)
    check5_pass = True
    check5_detail = []
    if in_progress_month:
        for w in in_progress_month.weeks:
            expected_elapsed_days = sum(1 for d in w.days if d.day <= run_date)
            actual_elapsed_days = sum(1 for d in w.days if d.has_data)
            if expected_elapsed_days != actual_elapsed_days:
                check5_pass = False
                check5_detail.append(f"{in_progress_month.name} {w.label}: expected {expected_elapsed_days} "
                                      f"elapsed days, workbook marked {actual_elapsed_days} as has_data")
        future_months = [m for m in months if m.idx > run_date.month]
        for m in future_months:
            any_has_data = any(d.has_data for w in m.weeks for d in w.days)
            if any_has_data:
                check5_pass = False
                check5_detail.append(f"{m.name} is a future month but has a day column marked has_data")
    all_pass &= check5_pass
    print(f"  [5] To-date vs completed logic: {'PASS' if check5_pass else 'FAIL'}")
    print(f"      In-progress month: {in_progress_month.name if in_progress_month else 'none'}; "
          f"in-progress week: "
          f"{next((w.label for w in (in_progress_month.weeks if in_progress_month else []) if not w.is_complete), 'none')}")
    for d in check5_detail:
        print(f"      - {d}")

    # --- QA 6: column count sanity check (per sheet -- Excel's 16,384
    # limit applies to each sheet individually, and there are now 3)
    check6_pass = True
    for sheet_name, data_cols in sheet_col_counts.items():
        sheet_total = data_cols + 1  # + the label column
        sheet_ok = sheet_total < 16384
        check6_pass &= sheet_ok
        print(f"  [6] Column count sanity check ({sheet_name}): {'PASS' if sheet_ok else 'FAIL'} "
              f"(total columns = {sheet_total}, limit = 16384)")
    all_pass &= check6_pass

    print(f"=== QA {'PASSED' if all_pass else 'FAILED'} ===\n")
    return all_pass


# =========================================================================
# main
# =========================================================================


def main() -> int:
    load_dotenv()
    try:
        access_token = os.environ.get("HUBSPOT_ACCESS_TOKEN", "")
        client = HubSpotClient(access_token)
        ref = resolve_reference_data(client)

        run_date = datetime.now(timezone.utc).date()

        print("Fetching contacts...")
        all_contacts = fetch_contacts(client, ref, run_date)
        print(f"  Fetched {len(all_contacts)} contacts")
        print("Fetching associated deals...")
        attach_deals(client, ref, all_contacts)
        total_deals = sum(len(c["deals"]) for c in all_contacts)
        print(f"  Fetched {total_deals} associated deals\n")

        filtered_contacts, filter_steps = apply_overall_filters(all_contacts, ref)

        rows = build_rows(filtered_contacts)
        print(f"Built row hierarchy: {len(rows)} rows "
              f"({sum(1 for r in rows if r['level'] == 1)} Level-1, "
              f"{sum(1 for r in rows if r['level'] == 2)} Level-2, "
              f"{sum(1 for r in rows if r['level'] == 3)} Level-3)\n")

        leaf_counts = compute_leaf_counts(rows, ref)

        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        print("Building workbook...")
        wb, sheet_cell_values, sheet_col_counts = build_workbook(rows, leaf_counts, ref, run_date)
        wb.save(OUTPUT_PATH)

        print("Recalculating formulas...")
        recalculate_workbook(OUTPUT_PATH, sheet_cell_values)

        qa_passed = run_qa(rows, leaf_counts, ref, run_date, filtered_contacts, all_contacts, filter_steps,
                            sheet_col_counts)

        print(f"Wrote {OUTPUT_PATH}")
        if not qa_passed:
            print("QA FAILED -- see details above.", file=sys.stderr)
            return 1
        return 0
    except HubSpotError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001 - top-level guard, per spec: clear message, non-zero exit
        print(f"UNEXPECTED ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
