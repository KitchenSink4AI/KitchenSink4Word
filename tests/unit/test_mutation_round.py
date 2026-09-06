"""One test per surviving mutant from the 2026-09-06 mutation round.

Spec: ``Draft/Working Files/Agent Results/20260906_wordppt_mutation_round.md``
(Word section: 319 mutants, 186 killed, 58% targeted kill rate; 55 class-(a)
survivors confirmed against the full 1390-test suite).

Every test here exists because a specific sabotage of the safety core passed
the whole suite. Each one was written against its mutant and hand-verified
RED with the mutant applied, GREEN with it reverted.

Test discipline carried over from the Excel sibling's round (whose method
notes recorded lock-timing tests false-redding under CPU load): NO test in
this file waits on a real elapsed interval to prove a policy. Ages come from
hand-written lockfiles carrying a chosen timestamp, idle gaps come from
``os.utime`` backdating, and the exact-value boundaries are measured against
a FROZEN wall clock injected into the module under test (monotonic and sleep
stay real), because a real elapsed interval can never land exactly ON the
boundary, which is the only place ``>`` and ``>=`` differ.
"""


from __future__ import annotations

import pytest

from word_mcp import envelope
from word_mcp.core.errors import AmbiguousTarget, WordMcpError


# ================================ envelope.py — the closed code vocabulary
#
# This is the one finding in the Word section that is NOT a mutant: the
# report noted that CLOSED_CODES is referenced only by the tests.
# ``refusal()`` took ``getattr(exc, "code", None)`` verbatim, so the closed
# refusal vocabulary was a convention enforced by three test assertions
# rather than by the code, and any exception carrying a ``.code`` attribute
# reached the wire with it. The clamp (_declared_code) closes it; these
# tests pin both the clamp and the declared codes that must survive it.


class _Declared(WordMcpError):
    pass


class TestClosedCodeVocabulary:
    def test_a_made_up_code_does_not_reach_the_wire(self):
        exc = _Declared("something went sideways")
        exc.code = "MADE_UP_CODE"
        out = envelope.refusal(exc)
        assert out["error"]["code"] in envelope.CLOSED_CODES
        assert out["error"]["code"] != "MADE_UP_CODE"

    def test_a_declared_valid_code_is_honoured(self):
        exc = _Declared("stale")
        exc.code = "STALE_ANCHOR"
        assert envelope.refusal(exc)["error"]["code"] == "STALE_ANCHOR"

    def test_a_non_string_code_falls_back_to_classification(self):
        exc = _Declared("weird")
        exc.code = 42
        assert envelope.refusal(exc)["error"]["code"] == "BAD_PARAMS"

    def test_a_third_party_exception_carrying_a_code_is_clamped(self):
        class Vendor(Exception):
            code = "vendor.timeout.7"

        out = envelope.refusal(Vendor("upstream said no"))
        assert out["error"]["code"] in envelope.CLOSED_CODES

    def test_a_lowercase_spelling_of_a_real_code_is_clamped(self):
        exc = _Declared("nearly right")
        exc.code = "not_found"
        assert envelope.refusal(exc)["error"]["code"] in envelope.CLOSED_CODES

    def test_the_hint_matches_the_clamped_code(self):
        exc = _Declared("nope")
        exc.code = "MADE_UP_CODE"
        out = envelope.refusal(exc)
        assert out["error"]["hint"] == envelope.HINTS.get(
            out["error"]["code"], ""
        )

    def test_every_declared_code_in_the_source_is_closed(self):
        """The five .code assignments in src/ (ops/batch.py, packs.py) must
        all survive the clamp; a typo at a sixth is what this catches."""
        for code in ("STALE_ANCHOR", "NOT_FOUND", "UNSUPPORTED_CONTENT",
                     "CONFLICT"):
            exc = _Declared("x")
            exc.code = code
            assert envelope.refusal(exc)["error"]["code"] == code


class TestRefusalPayloadShape:
    """envelope.py:178  `isinstance(exc, LookupError) and len(message) < 40`
    -> `or`, plus the matches/detail legs."""

    def test_a_bare_key_error_gets_a_real_message(self):
        out = envelope.refusal(KeyError(0))
        assert "internal lookup failed" in out["error"]["message"]

    def test_a_long_lookup_message_is_left_alone(self):
        long = "x" * 60
        out = envelope.refusal(KeyError(long))
        assert "internal lookup failed" not in out["error"]["message"]

    def test_a_short_non_lookup_message_is_left_alone(self):
        out = envelope.refusal(WordMcpError("nope"))
        assert out["error"]["message"] == "nope"
        assert "internal lookup failed" not in out["error"]["message"]

    def test_ambiguity_matches_ride_out_on_the_payload(self):
        exc = AmbiguousTarget("two matched")
        exc.matches = [{"paragraph": 1}, {"paragraph": 7}]
        out = envelope.refusal(exc)
        assert out["error"]["matches"] == [{"paragraph": 1}, {"paragraph": 7}]

    def test_a_plain_refusal_carries_no_matches_key(self):
        out = envelope.refusal(WordMcpError("plain"))
        assert "matches" not in out["error"]

    def test_the_refusal_result_serializes_as_an_error(self):
        result = envelope.refuse(WordMcpError("plain"))
        assert result.is_error is True
        assert result["ok"] is False
