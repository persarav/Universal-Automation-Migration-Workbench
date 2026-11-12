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
