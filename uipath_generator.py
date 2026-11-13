import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from engine import (
    analyze_source, build_intermediate_representation, detect_language,
    parse_xaml, parse_json_workflow, generate_fallback, redact_sensitive_text,
)
from skill_loader import parse_skill_markdown, parse_github_repo_url


def test_vba_procedure_detection():
    code = '''Attribute VB_Name = "ModuleA"\nOption Explicit\nPublic Sub RunJob(ByVal sourcePath As String)\n    If Len(sourcePath) = 0 Then Exit Sub\n    LoadData sourcePath\nEnd Sub\nPrivate Function LoadData(ByVal path As String) As Long\n    LoadData = 1\nEnd Function\n'''
    result = analyze_source({"filename":"sample.bas", "files":[{"filename":"sample.bas","text":code,"language":"VBA module"}]})
    names = [x["name"] for x in result["symbols"]]
    assert "RunJob" in names and "LoadData" in names
    assert result["symbol_count"] == 2


def test_python_ast_inventory():
    code = 'class Worker:\n    def run(self, item):\n        return clean(item)\n\ndef clean(value):\n    return value.strip()\n'
    result = analyze_source({"filename":"sample.py", "files":[{"filename":"sample.py","text":code,"language":"Python"}]})
    assert {x["name"] for x in result["symbols"]} == {"Worker", "run", "clean"}


def test_xaml_activity_parse():
    xml = '''<Activity xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"><Sequence DisplayName="Main"><If DisplayName="Check"><WriteLine Text="ok" /></If></Sequence></Activity>'''
    data = parse_xaml(xml, "Main.xaml")
    assert any(x["type"] == "Sequence" for x in data["activities"])
    assert any(x["type"] == "If" for x in data["activities"])


def test_power_automate_json_parse():
    data = parse_json_workflow(json.dumps({"definition":{"actions":{"Compose":{"type":"Compose","inputs":"hi"}}}}))
    assert any(x["name"] == "Compose" for x in data["actions"])


def test_ir_and_fallback_are_traceable():
    result = analyze_source({"filename":"sample.py", "files":[{"filename":"sample.py","text":"def run():\n    pass\n","language":"Python"}]})
    ir = build_intermediate_representation(result)
    output = generate_fallback("Python", ir, result)
    assert ir["workflow"]["steps"][0]["name"] == "run"
    assert output["generated_files"]
    assert output["confidence"] == "low"


def test_agent_skill_frontmatter():
    s = parse_skill_markdown('---\nname: sample-skill\ndescription: Sample instructions\n---\n# Rules')
    assert s["name"] == "sample-skill"
    assert "# Rules" in s["content"]


def test_github_tree_subpath():
    assert parse_github_repo_url("https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate") == ("microsoft", "power-platform-skills", "main", "plugins/power-automate")


def test_uipath_xaml_fallback_is_well_formed_xml():
    import xml.etree.ElementTree as ET
    result = analyze_source({"filename":"sample.py", "files":[{"filename":"sample.py","text":"def run():\n    pass\n","language":"Python"}]})
    ir = build_intermediate_representation(result)
    output = generate_fallback("UiPath (XAML)", ir, result)
    xaml = next(x["content"] for x in output["generated_files"] if x["path"].endswith(".xaml"))
    ET.fromstring(xaml)


def test_blue_prism_fallback_is_explicit_blueprint():
    result = analyze_source({"filename":"sample.bas", "files":[{"filename":"sample.bas","text":"Public Sub RunIt()\nEnd Sub\n","language":"VBA module"}]})
    ir = build_intermediate_representation(result)
    output = generate_fallback("Blue Prism", ir, result)
    design = next(x["content"] for x in output["generated_files"] if x["path"].endswith(".json"))
    assert json.loads(design)["not_import_ready"] is True


def test_secret_scrubbing_before_model_transfer():
    safe = redact_sensitive_text('api_key = "abcdef123456"\nAuthorization: Bearer abcdef1234567890token')
    assert "abcdef123456" not in safe
    assert "[REDACTED]" in safe


def test_migration_result_normalizer_is_safe_for_invalid_generator_output():
    from engine import normalize_migration_result
    normalized = normalize_migration_result(None, "UiPath (XAML)", {"symbols": [{"name": "RunJob"}]})
    assert isinstance(normalized, dict)
    assert normalized["mapping_table"] == []
    assert normalized["generated_files"] == []
    assert normalized["source_logic"]["detected_procedures"] == ["RunJob"]


def test_migration_result_normalizer_sanitizes_generated_file_payloads():
    from engine import normalize_migration_result
    normalized = normalize_migration_result({
        "generated_files": [
            {"path": "a.xaml", "content": None},
            {"path": "b.json", "content": {"hello": "world"}},
            {"path": None, "content": "ignore"},
        ],
        "mapping_table": [None, {"source": "Run", "target": "Main.xaml"}],
    }, "UiPath (XAML)", {"symbols": []})
    assert normalized["generated_files"][0]["content"] == ""
    assert '"hello": "world"' in normalized["generated_files"][1]["content"]
    assert len(normalized["generated_files"]) == 2
    assert normalized["mapping_table"] == [{"source": "Run", "target": "Main.xaml"}]


def test_bundled_skills_are_loaded_offline():
    from skill_loader import load_bundled_skills
    skills = load_bundled_skills()
    names = {item["name"].lower() for item in skills}
    assert len(skills) >= 7
    assert any("uipath" in name for name in names)
    assert any("power" in name for name in names)
    assert any("vba" in name for name in names)


def test_uipath_generator_maps_sync_source_to_modular_project():
    import xml.etree.ElementTree as ET
    from uipath_generator import generate_uipath_project
    symbols = [
        {"name": "Run_Full_Club_Check", "kind": "Sub", "source_file": "AutoClubCheck.bas", "line": 1, "end_line": 10, "signature": "Sub Run_Full_Club_Check()", "calls": []},
        {"name": "STO_Check", "kind": "Sub", "source_file": "Module3.bas", "line": 1, "end_line": 10, "signature": "Sub STO_Check()", "calls": []},
        {"name": "PO_Check", "kind": "Sub", "source_file": "Module6.bas", "line": 1, "end_line": 10, "signature": "Sub PO_Check()", "calls": []},
        {"name": "Storage_Release_Check", "kind": "Sub", "source_file": "StorageRelease.bas", "line": 1, "end_line": 10, "signature": "Sub Storage_Release_Check()", "calls": []},
    ]
    analysis = {"symbols": symbols, "patterns": {}, "external_touchpoints": []}
    output = generate_uipath_project(analysis, {}, [])
    xamls = [f for f in output["generated_files"] if f["path"].endswith(".xaml")]
    assert len(xamls) == 20
    for item in xamls:
        ET.fromstring(item["content"])
    by_path = {f["path"]: f["content"] for f in xamls}
    replacement = by_path["Workflows/Replacement_Article.xaml"]
    assert "Order Quantity" in replacement
    assert "Still to be delivered (qty)" in replacement
    assert "newClubNo" in replacement
    # Run_All must preserve the VBA caller order and should not invoke unrelated
    # button-driven workflows such as Replacement Article.
    main = by_path["Main.xaml"]
    assert "Workflows/STO_Check.xaml" in main
    assert "Workflows/PO_Check.xaml" in main
    assert "Workflows/Storage_Release.xaml" in main
    assert "Workflows/Replacement_Article.xaml" not in main
    # Each user-triggered capability still has its own callable XAML workflow.
    for required in ["Workflows/Clear_Data.xaml", "Workflows/Email_Analysts.xaml", "Workflows/Appointment_Date.xaml", "Workflows/Fetch_STF_Article.xaml", "Workflows/Take_Daily_Snapshot.xaml"]:
        assert required in by_path
    assert output["target_design"]["uses_api_key"] is False


def test_sync_procedure_classification_keeps_helpers_with_owner_workflow():
    from uipath_generator import _classify_procedure
    cases = [
        ({"name":"TakeDailySnapshot", "source_file":"Module8.bas"}, "Workflows/Take_Daily_Snapshot.xaml"),
        ({"name":"BuildBestPO", "source_file":"AutoClubCheck.bas"}, "Workflows/PO_Check.xaml"),
        ({"name":"WritePartialComments", "source_file":"AutoClubCheck.bas"}, "Workflows/PO_Check.xaml"),
        ({"name":"BestReplacements", "source_file":"Module6.bas"}, "Workflows/Fetch_STF_Article.xaml"),
        ({"name":"HasStoOrPo", "source_file":"AutoClubCheck.bas"}, "Workflows/Appointment_Date.xaml"),
        ({"name":"SumChart", "source_file":"Module8.bas"}, "Workflows/UiOnly_Replacement.xaml"),
        ({"name":"ReadInvstat", "source_file":"AutoClubCheck.bas"}, "SAP_Adapters/Capture_SAP_UI.xaml"),
    ]
    for symbol, expected in cases:
        assert _classify_procedure(symbol)[0] == expected


def test_uipath_invoke_code_arguments_bind_to_variable_names():
    from uipath_generator import workflow_xaml
    xaml = workflow_xaml(
        "TestWorkflow", "Test", "test map",
        [("in_value", "In", "x:String"), ("out_updated", "Out", "x:Int32")],
        'outUpdated = inValue.Length'
    )
    assert 'x:Key="value"' in xaml
    assert 'x:Key="outUpdated"' in xaml
    assert '[out_updated]' in xaml


def test_sync_uipath_generator_emits_studio_24108_type_and_activity_namespaces():
    import json
    import xml.etree.ElementTree as ET
    from uipath_generator import generate_uipath_project
    symbols = [
        {"name":"Run_Full_Club_Check", "kind":"Sub", "source_file":"AutoClubCheck.bas", "line":1, "end_line":10, "signature":"Sub Run_Full_Club_Check()", "calls":[]},
        {"name":"STO_Check", "kind":"Sub", "source_file":"Module3.bas", "line":1, "end_line":10, "signature":"Sub STO_Check()", "calls":[]},
        {"name":"PO_Check", "kind":"Sub", "source_file":"Module6.bas", "line":1, "end_line":10, "signature":"Sub PO_Check()", "calls":[]},
        {"name":"Storage_Release_Check", "kind":"Sub", "source_file":"StorageRelease.bas", "line":1, "end_line":10, "signature":"Sub Storage_Release_Check()", "calls":[]},
    ]
    output = generate_uipath_project({"symbols":symbols, "patterns":{}, "external_touchpoints":[]}, {}, [])
    by_path = {f["path"]:f["content"] for f in output["generated_files"]}
    project = json.loads(by_path["project.json"])
    assert project["studioVersion"] == "24.10.8"
    xamls = {k:v for k,v in by_path.items() if k.endswith(".xaml")}
    assert len(xamls) == 20
    for path, xml in xamls.items():
        ET.fromstring(xml)
        assert "x:DateTime" not in xml, path
        assert "<ue:" not in xml and "xmlns:ue=" not in xml, path
        if "sys:DateTime" in xml:
            assert 'xmlns:sys="clr-namespace:System;assembly=System.Private.CoreLib"' in xml, path
    main = xamls["Main.xaml"]
    assert "<ui:ReadRange" in main
    assert "<ui:WriteRange" in main
    assert "xmlns:ui=\"http://schemas.uipath.com/workflow/activities\"" in main
    assert "xmlns:sys=\"clr-namespace:System;assembly=System.Private.CoreLib\"" in main
