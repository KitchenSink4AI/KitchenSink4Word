r"""Issue #24: get_server_info called Word "not registered" on a machine
that had Word.

The reporter ran KitchenSink4Word 2.1.2 on Windows 11 with Word 365
installed, registered under HKCR\Word.Application and automating fine,
and got back application "not registered", com_tools "unavailable". The
cause was one line in the registration probe:

    pythoncom.CLSIDFromProgID("Word.Application")

pywin32's pythoncom has ProgIDFromCLSID. It has never had
CLSIDFromProgID. The call raised AttributeError on every machine, a
broad ``except Exception`` caught it, and the handler reported the one
conclusion it knew how to report: no Word here. The probe never reached
the registry, so the answer was the same whether Word was installed or
not.

These tests reproduce that shape without a real registry and without
COM, so they run on any machine. They fail against the 2.1.2 probe and
pass once the probe reads the registry and separates "I could not tell"
from "you have no Word".
"""

import sys
import types

import pytest

from word_mcp import server


_CLSID = "{000209FF-0000-0000-C000-000000000046}"
_PROGID_KEY = "Word.Application" + "\\" + "CLSID"
_SERVER_KEY = "CLSID" + "\\" + _CLSID + "\\" + "LocalServer32"


class _Key:
    def __init__(self, value=""):
        self._value = value

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class _RegistryWithWord:
    """A registry that answers the way the reporter's machine answers."""

    HKEY_CLASSES_ROOT = 0
    KEY_READ = 0x20019
    KEY_WOW64_64KEY = 0x0100
    KEY_WOW64_32KEY = 0x0200

    keys = {
        _PROGID_KEY: _CLSID,
        _SERVER_KEY: (
            r"C:\Program Files\Microsoft Office\root\Office16"
            r"\WINWORD.EXE /Automation"
        ),
    }

    def OpenKey(self, root, sub, reserved=0, access=0):
        if sub in self.keys:
            return _Key(self.keys[sub])
        raise FileNotFoundError(2, "no such key")

    def QueryValueEx(self, key, name):
        return (key._value, 1)


class _RegistryThatBlowsUp:
    """A registry read that fails in a way nobody planned for."""

    HKEY_CLASSES_ROOT = 0
    KEY_READ = 0x20019
    KEY_WOW64_64KEY = 0x0100
    KEY_WOW64_32KEY = 0x0200

    def OpenKey(self, root, sub, reserved=0, access=0):
        raise AttributeError(
            "module 'pythoncom' has no attribute 'CLSIDFromProgID'"
        )

    def QueryValueEx(self, key, name):  # pragma: no cover - never reached
        raise AssertionError("the open should have failed first")


def _pywin32_as_it_actually_ships():
    """pythoncom as pywin32 publishes it: no CLSIDFromProgID on it."""
    mod = types.ModuleType("pythoncom")
    mod.ProgIDFromCLSID = lambda clsid: "Word.Application"
    return mod


@pytest.fixture
def reporters_machine(monkeypatch):
    """Windows, pywin32 installed, Word registered. Nothing else stubbed."""
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(
        sys.modules, "pythoncom", _pywin32_as_it_actually_ships()
    )
    monkeypatch.setitem(sys.modules, "winreg", _RegistryWithWord())


def test_installed_word_is_not_reported_as_not_registered(reporters_machine):
    """The report the reporter got, and must not get again."""
    word = server._word_environment()
    assert word["pywin32"] == "installed"
    assert word["application"] != "not registered", (
        "issue #24: Word is registered on this machine and the report "
        "said it was not"
    )
    assert word["application"] == "installed"
    assert word["com_tools"] == "available"


def test_a_probe_failure_is_reported_rather_than_hidden(monkeypatch):
    """The other half of #24: the failure was swallowed.

    An AttributeError raised inside the probe used to come back as a
    confident statement about the machine. Whatever goes wrong in there,
    the report has to say the probe could not tell, and it may not claim
    the machine has no Word.
    """
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setitem(
        sys.modules, "pythoncom", _pywin32_as_it_actually_ships()
    )
    monkeypatch.setitem(sys.modules, "winreg", _RegistryThatBlowsUp())
    word = server._word_environment()
    assert word["application"] != "not registered"
    assert word["com_tools"] != "available"
    note = word["note"].lower()
    assert "attributeerror" in note, (
        "the report has to name what went wrong, not hide it"
    )
    assert "no word.application registration found" not in note


def test_the_attribute_the_old_probe_called_does_not_exist():
    """Straight from the report, checked against the installed pywin32."""
    pythoncom = pytest.importorskip("pythoncom")
    assert not hasattr(pythoncom, "CLSIDFromProgID")
    assert hasattr(pythoncom, "ProgIDFromCLSID")


def test_the_progid_resolves_through_the_api_that_does_exist():
    """The reporter's own suggestion, run here: pywintypes.IID resolves
    Word.Application to the class id the registry holds."""
    pywintypes = pytest.importorskip("pywintypes")
    try:
        iid = pywintypes.IID("Word.Application")
    except Exception:  # pragma: no cover - a machine with no Word
        pytest.skip("no Word.Application registration on this machine")
    assert str(iid).upper() == _CLSID


def test_the_source_no_longer_calls_the_missing_attribute():
    """A grep guard, so the line cannot come back by a different route."""
    import inspect

    src = inspect.getsource(server)
    assert "CLSIDFromProgID" not in src
