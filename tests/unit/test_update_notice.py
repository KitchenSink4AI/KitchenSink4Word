"""The on-demand update notice: one surface, one call a day, honest output.

Nothing here touches the network. Every test that would reach PyPI patches
the fetch, and the tests that prove a path does NOT reach the network patch
it with a counter or a landmine.

The properties pinned:

1. THE NOTICE APPEARS. A newer published stable produces state
   'update_available', both version numbers, and the pinned-install fact.
2. HONEST FAILURE. A check that cannot reach PyPI reports that it could
   not, with the reason and the age of the last answer that did arrive.
   It is never silently absent and it never invents a version.
3. ONE CALL A WEEK. A second call inside seven days reaches the cache,
   not the network; past seven days it asks again.
4. THE OFF SWITCH. KS4W_UPDATE_CHECK=off produces a report saying the
   check is off, and performs no network, no cache read, no cache write.
5. HOSTILE INPUT. Malformed, wrong-typed, and adversarial PyPI payloads
   never crash and never fabricate a version.
6. ONE SURFACE. get_server_info is the only tool that can reach the
   network. No other tool path fires the check.
7. PRIVACY. The first report discloses that the check is a plain HTTPS GET
   to pypi.org, and the request carries no body and no identifier.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from word_mcp import __version__, server
from word_mcp.core import update_check as uc


@pytest.fixture(autouse=True)
def _isolated(tmp_path, monkeypatch):
    """Every test gets its own cache directory and the check switched on."""
    monkeypatch.delenv(uc.OFF_ENV, raising=False)
    monkeypatch.delenv(uc.LEGACY_OFF_ENV, raising=False)
    monkeypatch.setenv(uc.CACHE_DIR_ENV, str(tmp_path / "state"))
    return tmp_path


def _payload(*versions):
    releases = {v: [{"filename": f"x-{v}.whl", "yanked": False}]
                for v in versions}
    return {"info": {"version": versions[-1] if versions else ""},
            "releases": releases}


class _Counter:
    """A fetch stand-in that counts how often the network would be used."""

    def __init__(self, payload=None, raises=None):
        self.payload = payload
        self.raises = raises
        self.calls = 0

    def __call__(self, *args, **kwargs):
        self.calls += 1
        if self.raises is not None:
            raise self.raises
        return self.payload


def _landmine(*args, **kwargs):
    raise AssertionError("this path must not touch the network or the cache")


def _write_cache(latest, *, ok=True, age_hours=0.0, **extra):
    stamp = datetime.now(timezone.utc) - timedelta(hours=age_hours)
    uc.write_cache({"last_check": stamp.isoformat(),
                    "latest_version": latest, "ok": ok,
                    "last_success": stamp.isoformat() if ok else None,
                    **extra})


# ------------------------------------------------------- 1. the notice


def test_a_newer_release_produces_the_notice(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(_payload("999.0.0")))
    out = uc.status()
    assert out["state"] == "update_available"
    assert out["current_version"] == __version__
    assert out["latest_version"] == "999.0.0"
    assert out["install_note"]
    assert out["reachable"] is True
    assert out["last_successful_check"]


def test_the_running_build_being_current_says_so(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(_payload(__version__)))
    out = uc.status()
    assert out["state"] == "current"
    assert out["latest_version"] == __version__
    assert "install_note" not in out


def test_get_server_info_carries_the_report(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(_payload("999.0.0")))
    info = server.get_server_info()
    assert info["update_check"]["state"] == "update_available"
    assert info["update_check"]["latest_version"] == "999.0.0"


# ---------------------------------------------------- 2. honest failure


@pytest.mark.parametrize("blow_up", [
    TimeoutError("timed out"),
    OSError("network unreachable"),
    ValueError("not json"),
])
def test_an_unreachable_index_is_reported_not_hidden(blow_up, monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(raises=blow_up))
    out = uc.status()
    assert out["reachable"] is False
    assert type(blow_up).__name__ in out["error"]
    assert out["note"]                      # never silently absent
    assert "latest_version" not in out      # and never invented
    assert out["state"] == "unknown"


def test_a_failure_reports_the_age_of_the_last_good_answer(monkeypatch):
    _write_cache("999.0.0", ok=True, age_hours=24.0 * 8)
    stamp = uc.read_cache()["last_success"]
    monkeypatch.setattr(uc, "_fetch", _Counter(raises=OSError("down")))
    out = uc.status()
    assert out["reachable"] is False
    assert out["last_successful_check"] == stamp
    assert out["latest_version"] == "999.0.0"   # the answer that did arrive
    assert out["state"] == "update_available"


def test_a_never_successful_check_admits_it_knows_nothing(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(raises=OSError("down")))
    out = uc.status()
    assert out["last_successful_check"] is None
    assert out["state"] == "unknown"


# ------------------------------------------------------ 3. one call a week


def test_a_second_call_inside_the_window_skips_the_network(monkeypatch):
    counter = _Counter(_payload("999.0.0"))
    monkeypatch.setattr(uc, "_fetch", counter)
    first = uc.status()
    second = uc.status()
    assert counter.calls == 1
    assert second["latest_version"] == first["latest_version"] == "999.0.0"


def test_a_failed_attempt_also_holds_the_network_off_for_the_window(
        monkeypatch):
    counter = _Counter(raises=OSError("down"))
    monkeypatch.setattr(uc, "_fetch", counter)
    uc.status()
    uc.status()
    assert counter.calls == 1


def test_past_the_window_the_check_asks_again(monkeypatch):
    _write_cache("1.0.0", ok=True, age_hours=24.0 * 8)
    counter = _Counter(_payload("999.0.0"))
    monkeypatch.setattr(uc, "_fetch", counter)
    out = uc.status()
    assert counter.calls == 1
    assert out["latest_version"] == "999.0.0"


def test_a_days_old_answer_is_still_good_enough_to_skip_the_network(
        monkeypatch):
    """Three days in, the cached answer still serves. This is the cadence
    the products want: a release a week does not become a notice a week."""
    _write_cache("999.0.0", ok=True, age_hours=24.0 * 3)
    counter = _Counter(_payload("1000.0.0"))
    monkeypatch.setattr(uc, "_fetch", counter)
    out = uc.status()
    assert counter.calls == 0
    assert out["latest_version"] == "999.0.0"


def test_the_horizon_is_seven_days(monkeypatch):
    assert uc.CHECK_INTERVAL == timedelta(days=7)
    assert uc.CHECK_INTERVAL.total_seconds() == 604800
    fresh = datetime.now(timezone.utc) - timedelta(days=6)
    aged = datetime.now(timezone.utc) - timedelta(days=8)
    assert uc.is_due({"last_check": fresh.isoformat(), "ok": True}) is False
    assert uc.is_due({"last_check": aged.isoformat(), "ok": True}) is True


def test_the_timeout_is_two_seconds(monkeypatch):
    assert uc.TIMEOUT_SECONDS == 2.0
    seen = {}

    class _Resp:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self, *a):
            return b'{"info": {"version": "1.0.0"}}'

    def _urlopen(request, timeout=None):
        seen["timeout"] = timeout
        seen["url"] = request.full_url
        seen["body"] = request.data
        return _Resp()

    monkeypatch.setattr(uc.urllib.request, "urlopen", _urlopen)
    uc._fetch()
    assert seen["timeout"] == 2.0


# -------------------------------------------------------- 4. the off switch


@pytest.mark.parametrize("value", ["off", "OFF", " off "])
def test_off_switch_values(value, monkeypatch):
    monkeypatch.setenv(uc.OFF_ENV, value)
    assert uc.disabled() is True


@pytest.mark.parametrize("value", ["", "on", "1", "true"])
def test_non_off_values_leave_the_check_on(value, monkeypatch):
    monkeypatch.setenv(uc.OFF_ENV, value)
    assert uc.disabled() is False


def test_off_says_so_and_does_no_io(monkeypatch):
    monkeypatch.setenv(uc.OFF_ENV, "off")
    monkeypatch.setattr(uc, "_fetch", _landmine)
    monkeypatch.setattr(uc, "read_cache", _landmine)
    monkeypatch.setattr(uc, "write_cache", _landmine)
    out = uc.status()
    assert out["state"] == "disabled"
    assert out["disabled_by"] == uc.OFF_ENV
    assert out["off_value"] == "off"
    assert out["note"]


def test_off_is_visible_through_the_server_surface(monkeypatch):
    monkeypatch.setenv(uc.OFF_ENV, "off")
    monkeypatch.setattr(uc, "_fetch", _landmine)
    assert server.get_server_info()["update_check"]["state"] == "disabled"


def test_the_legacy_opt_out_still_disables(monkeypatch):
    monkeypatch.setenv(uc.LEGACY_OFF_ENV, "1")
    assert uc.disabled() is True
    monkeypatch.setattr(uc, "_fetch", _landmine)
    assert uc.status()["state"] == "disabled"


# ------------------------------------------------------- 5. hostile input


@pytest.mark.parametrize("payload", [
    {},
    {"info": {}},
    {"info": None},
    {"releases": "nonsense"},
    {"releases": {"'; DROP TABLE releases; --": [{"filename": "x"}]}},
    {"info": {"version": "not.a.version"}},
    {"info": {"version": ["9", "9"]}},
    {"info": {"version": {"nested": "object"}}},
    {"releases": {"9" * 5000: [{"filename": "x"}]}},
    [1, 2, 3],
    "a string, not an object",
    None,
])
def test_hostile_payloads_never_crash_and_never_invent_a_version(
        payload, monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(payload))
    out = uc.status()
    assert out["state"] == "unknown"
    assert "latest_version" not in out
    assert out["reachable"] is False
    assert out["current_version"] == __version__


def test_a_wrong_typed_release_map_reports_only_what_it_contained(
        monkeypatch):
    """A version key whose file list is the wrong type is still a version
    the index published; what must never happen is a number appearing that
    the payload did not contain."""
    monkeypatch.setattr(uc, "_fetch",
                        _Counter({"releases": {"1.0.0": "not a list"}}))
    out = uc.status()
    latest = out.get("latest_version")
    assert latest is None or latest == "1.0.0"


def test_a_version_shaped_lie_is_still_compared_not_trusted(monkeypatch):
    """An index claiming an ancient release is newest does not produce a
    notice; the comparison decides, not the payload."""
    monkeypatch.setattr(uc, "_fetch", _Counter(_payload("0.0.1")))
    out = uc.status()
    assert out["state"] == "current"


# --------------------------------------------------------- 6. one surface


def test_no_other_tool_path_fires_the_check(tmp_path, monkeypatch):
    counter = _Counter(_payload("999.0.0"))
    monkeypatch.setattr(uc, "_fetch", counter)
    doc = str(tmp_path / "d.docx")
    server.create_document(doc)
    server.get_document_info(doc)
    server.diagnose_document(doc)
    server.get_outline(doc)
    server.get_workflows()
    server.get_workflows("bulk-edit")
    assert counter.calls == 0
    server.get_server_info()
    assert counter.calls == 1


def test_nothing_starts_a_thread(monkeypatch):
    """The check runs on demand: the module offers no background starter."""
    assert not hasattr(uc, "start_background_check")
    assert "threading" not in dir(uc)


# ----------------------------------------------------------- 7. privacy


def test_the_first_report_discloses_the_request(monkeypatch):
    monkeypatch.setattr(uc, "_fetch", _Counter(_payload("999.0.0")))
    first = uc.status()
    assert first["privacy"]
    assert first["privacy_endpoint"] == "https://pypi.org/pypi/" \
        "kitchensink4word/json"
    second = uc.status()
    assert "privacy" not in second


def test_the_request_carries_no_body_and_no_identifier(monkeypatch):
    seen = {}

    class _Resp:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self, *a):
            return b'{"info": {"version": "1.0.0"}}'

    def _urlopen(request, timeout=None):
        seen["url"] = request.full_url
        seen["data"] = request.data
        seen["method"] = request.get_method()
        seen["headers"] = dict(request.header_items())
        return _Resp()

    monkeypatch.setattr(uc.urllib.request, "urlopen", _urlopen)
    uc._fetch()
    assert seen["url"].startswith("https://pypi.org/")
    assert seen["data"] is None
    assert seen["method"] == "GET"
    values = " ".join(str(v) for v in seen["headers"].values())
    assert "kitchensink4word" in values          # package and version only
    assert "Cookie" not in seen["headers"]
    assert "Authorization" not in seen["headers"]
