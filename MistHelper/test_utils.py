import sys
import pytest
import mistapi  # <-- Add this import
from MistHelper import MistHelper

def test_cli_menu_action_valid(monkeypatch):
    # Simulate CLI args for a valid menu action (e.g., "11" for export_org_site_list)
    test_args = ["MistHelper.py", "-M", "11"]
    monkeypatch.setattr(sys, "argv", test_args)
    # Patch the menu action function to track if it was called
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])  # <-- Patch here
    # Run main CLI block
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert called.get("ran") is True
    assert e.value.code == 0

def test_cli_menu_action_invalid(monkeypatch, capsys):
    # Simulate CLI args for an invalid menu action
    test_args = ["MistHelper.py", "-M", "99"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    captured = capsys.readouterr()
    assert "Invalid menu option" in captured.out
    assert e.value.code == 1

def test_cli_site_name_resolution(monkeypatch):
    # Simulate CLI args with a site name that does not exist
    test_args = ["MistHelper.py", "-M", "11", "-S", "NonexistentSite"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    # Patch MistHelper.mistapi.get_all to return empty list for sites
    monkeypatch.setattr(MistHelper.mistapi, "get_all", lambda *a, **kw: [])
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert e.value.code == 1

def test_interactive_menu(monkeypatch, capsys):
    # Simulate no CLI args and user entering a valid menu option
    test_args = ["MistHelper.py"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr("builtins.input", lambda _: "11")
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    with pytest.raises(SystemExit):
        MistHelper.main()
    assert called.get("ran") is True

def test_interactive_menu_invalid(monkeypatch, capsys):
    # Simulate no CLI args and user entering an invalid menu option
    test_args = ["MistHelper.py"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr("builtins.input", lambda _: "99")
    with pytest.raises(SystemExit):
        MistHelper.main()
    captured = capsys.readouterr()
    assert "Invalid selection" in captured.out

from MistHelper.MistHelper import (
    flatten_nested_dict,
    flatten_all_nested_fields,
    escape_multiline_strings,
    convert_list_values_to_strings,
    get_all_unique_keys,
)

def test_flatten_nested_dict_simple():
    d = {"a": 1, "b": {"c": 2, "d": 3}}
    flat = flatten_nested_dict(d)
    assert flat == {"a": 1, "b_c": 2, "b_d": 3}

def test_flatten_nested_dict_with_list():
    d = {"a": [1, 2, 3], "b": {"c": [4, 5]}}
    flat = flatten_nested_dict(d)
    assert flat["a"] == "1,2,3"
    assert flat["b_c"] == "4,5"

def test_flatten_nested_dict_with_list_of_dicts():
    d = {"a": [{"x": 1}, {"y": 2}], "b": 3}
    flat = flatten_nested_dict(d)
    assert flat["a_0_x"] == 1
    assert flat["a_1_y"] == 2
    assert flat["b"] == 3

def test_flatten_all_nested_fields():
    data = [
        {"a": 1, "b": {"c": 2}},
        {"a": 2, "b": {"c": 3, "d": 4}}
    ]
    flat = flatten_all_nested_fields(data)
    assert flat[0]["b_c"] == 2
    assert flat[1]["b_d"] == 4

def test_flatten_all_nested_fields_with_stringified_dict():
    data = [{"a": "{'x': 1, 'y': 2}"}]
    flat = flatten_all_nested_fields(data)
    # Should flatten stringified dict under key 'a'
    assert "a_x" in flat[0]
    assert "a_y" in flat[0]
    assert flat[0]["a_x"] == 1
    assert flat[0]["a_y"] == 2

def test_escape_multiline_strings():
    data = [{"a": "hello\nworld", "b": ["x", "y"]}]
    escaped = escape_multiline_strings(data)
    assert escaped[0]["a"] == "hello\\nworld"
    assert escaped[0]["b"] == "x,y"

def test_escape_multiline_strings_with_carriage_return():
    data = [{"a": "line1\r\nline2"}]
    escaped = escape_multiline_strings(data)
    assert escaped[0]["a"] == "line1\\nline2"

def test_convert_list_values_to_strings():
    data = [{"a": [1, 2, 3], "b": "test"}]
    converted = convert_list_values_to_strings(data)
    assert converted[0]["a"] == "1,2,3"
    assert converted[0]["b"] == "test"

def test_get_all_unique_keys():
    data = [{"a": 1, "b": 2}, {"b": 3, "c": 4}]
    keys = get_all_unique_keys(data)
    assert set(keys) == {"a", "b", "c"}

def test_flatten_nested_dict_empty():
    d = {}
    flat = flatten_nested_dict(d)
    assert flat == {}

def test_flatten_all_nested_fields_empty():
    data = []
    flat = flatten_all_nested_fields(data)
    assert flat == []

def test_escape_multiline_strings_empty():
    data = []
    escaped = escape_multiline_strings(data)
    assert escaped == []

def test_main_exits_properly(monkeypatch):
    test_args = ["MistHelper.py"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr("builtins.input", lambda _: "11")  # or any valid menu option
    # Patch the menu action function to avoid side effects
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert called.get("ran") is True
    assert e.value.code == 0

def test_cli_with_org_arg(monkeypatch):
    # Simulate CLI args with org argument and valid menu
    test_args = ["MistHelper.py", "-O", "test-org", "-M", "11"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["test-org"])
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert called.get("ran") is True
    assert e.value.code == 0

def test_cli_with_invalid_site(monkeypatch, capsys):
    # Simulate CLI args with a site name that does not exist
    test_args = ["MistHelper.py", "-M", "11", "-S", "FakeSite"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    monkeypatch.setattr(mistapi, "get_all", lambda *a, **kw: [{"name": "RealSite", "id": "123"}])
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    captured = capsys.readouterr()
    assert "not found" in captured.out
    assert e.value.code == 1

def test_cli_with_invalid_device(monkeypatch, capsys):
    # Simulate CLI args with a valid site but invalid device name
    test_args = ["MistHelper.py", "-M", "11", "-S", "SiteA", "-D", "NoDevice"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    monkeypatch.setattr(
        mistapi, "get_all",
        lambda *a, **kw: [{"name": "SiteA", "id": "site-1"}] if "sites" in str(a[0]) else [{"name": "DeviceA", "id": "dev-1"}]
    )
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    captured = capsys.readouterr()
    assert "not found" in captured.out
    assert e.value.code == 1

def test_cli_with_valid_site_and_device(monkeypatch):
    # Simulate CLI args with valid site and device names
    test_args = ["MistHelper.py", "-M", "11", "-S", "SiteA", "-D", "DeviceA"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    # Always return the site for any get_all call that looks for sites
    def fake_get_all(*a, **kw):
        if "site" in str(a).lower():
            return [{"name": "SiteA", "id": "site-1"}]
        if "device" in str(a).lower():
            return [{"name": "DeviceA", "id": "dev-1"}]
        return []
    monkeypatch.setattr(mistapi, "get_all", fake_get_all)
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert called.get("ran") is True
    assert e.value.code == 0

def test_interactive_menu_empty_input(monkeypatch, capsys):
    # Simulate user pressing enter (empty input) in interactive menu
    test_args = ["MistHelper.py"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr("builtins.input", lambda _: "")
    with pytest.raises(SystemExit):
        MistHelper.main()
    captured = capsys.readouterr()
    assert "Invalid selection" in captured.out

def test_interactive_menu_non_numeric_input(monkeypatch, capsys):
    # Simulate user entering a non-numeric invalid menu option
    test_args = ["MistHelper.py"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr("builtins.input", lambda _: "foobar")
    with pytest.raises(SystemExit):
        MistHelper.main()
    captured = capsys.readouterr()
    assert "Invalid selection" in captured.out

def test_cli_with_port_arg(monkeypatch):
    # Simulate CLI args with port argument (should not affect menu action)
    test_args = ["MistHelper.py", "-M", "11", "-P", "eth0"]
    monkeypatch.setattr(sys, "argv", test_args)
    monkeypatch.setattr(mistapi.cli, "select_org", lambda apisession: ["dummy-org-id"])
    called = {}
    def fake_export_org_site_list():
        called["ran"] = True
    MistHelper.menu_actions["11"] = (fake_export_org_site_list, "desc")
    with pytest.raises(SystemExit) as e:
        MistHelper.main()
    assert called.get("ran") is True
    assert e.value.code == 0