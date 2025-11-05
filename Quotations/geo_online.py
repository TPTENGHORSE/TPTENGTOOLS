import os
from typing import Optional, Tuple

try:
    from geopy.geocoders import Nominatim  # type: ignore
    from geopy.extra.rate_limiter import RateLimiter  # type: ignore
except Exception:  # pragma: no cover
    Nominatim = None  # type: ignore
    RateLimiter = None  # type: ignore


def _allow_online_for_country(country_code: str) -> bool:
    """Determine whether online geocoding is allowed.
    - Default: allow only for CN (China) to solve coverage gaps.
    - Override with env var QTOOL_ONLINE_GEOCODING:
      * '0'/'false'/'no' -> disable
      * '1'/'true'/'yes' -> enable for all countries
      * 'cn' -> enable only for CN (default behavior)
    """
    raw = (os.environ.get("QTOOL_ONLINE_GEOCODING") or "cn").strip().lower()
    if raw in ("0", "false", "no"):
        return False
    if raw in ("1", "true", "yes"):
        return True
    # default or 'cn'
    return (country_code or "").strip().upper() == "CN"


def geocode_city_nominatim(city: str, country_code_or_name: str, timeout: int = 10) -> Tuple[Optional[float], Optional[float], str]:
    """Try to geocode a city using Nominatim (OpenStreetMap) with a modest rate limit.
    Returns (lat, lon, source_tag) or (None, None, '').
    Notes:
    - Requires network connectivity and respects Nominatim usage policy (custom user agent, limited rate).
    - For country scoping, we pass `country_codes` when the provided country looks like ISO2.
    """
    if Nominatim is None or RateLimiter is None:
        return None, None, ""
    if not city or not country_code_or_name:
        return None, None, ""
    cc = (country_code_or_name or "").strip()
    if not _allow_online_for_country(cc):
        return None, None, ""
    try:
        # Build user agent. If you can, set QTOOL_GEO_UA to your email or service URL for better compliance.
        ua = os.environ.get("QTOOL_GEO_UA", "horse-qtool/1.0 (nominatim,fallback)")
        geolocator = Nominatim(user_agent=ua, timeout=timeout)
        rate_limited = RateLimiter(geolocator.geocode, min_delay_seconds=1.0)
        # Use country_codes when we have ISO2; otherwise, include country in the query string
        cc_upper = cc.strip().upper()
        is_iso2 = len(cc_upper) == 2 and cc_upper.isalpha()
        query = f"{city}, {country_code_or_name}"
        if is_iso2:
            location = rate_limited(query, exactly_one=True, addressdetails=False, language="en", country_codes=cc_upper.lower())
        else:
            location = rate_limited(query, exactly_one=True, addressdetails=False, language="en")
        if location is None:
            return None, None, ""
        return float(location.latitude), float(location.longitude), "nominatim"
    except Exception:
        return None, None, ""


def geocode_city_online_if_allowed(country_code: str, city: str) -> Tuple[Optional[float], Optional[float], str]:
    """Convenience wrapper that checks allowlist and returns coordinates if available."""
    lat, lon, src = geocode_city_nominatim(city, country_code)
    if lat is not None and lon is not None:
        return lat, lon, src
    return None, None, ""
