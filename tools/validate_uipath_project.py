#!/usr/bin/env python3
"""Static structural validation for the SYNC UiPath project.

This script does not replace UiPath Studio's XAML serializer, expression compiler,
package restore, Validate, or Build. Run validate_all.ps1 on Windows with UiPath CLI
for the authoritative platform gate.
"""
from __future__ import annotations
import argparse, csv, json, re
from pathlib import Path
import xml.etree.ElementTree as ET

ACT = "http://schemas.microsoft.com/netfx/2009/xaml/activities"
X = "http://schemas.microsoft.com/winfx/2006/xaml"
UI = "http://schemas.uipath.com/workflow/activities"
SAP2010 = "http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation"
SYS_CLR = "clr-namespace:System;assembly=System.Private.CoreLib"
SCG_CLR = "clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib"
XKEY = f"{{{X}}}Key"
XCLASS = f"{{{X}}}Class"
IDREF = f"{{{SAP2010}}}WorkflowViewState.IdRef"

X_BUILTINS = {"String", "Int32", "Int64", "Double", "Boolean", "Byte", "Single", "Decimal", "Char", "Object", "TimeSpan"}
VALID_ARG_WRAPPERS = {"InArgument", "OutArgument", "InOutArgument"}


def namespaces(path: Path) -> dict[str, str]:
    found: dict[str, str] = {}
    for _event, item in ET.iterparse(path, events=("start-ns",)):
        prefix, uri = item
        found[prefix or ""] = uri
    return found


def check_type(type_text: str, ns: dict[str, str], where: str, errors: list[str]) -> None:
    """Validate CLR type aliases in UiPath's XAML-friendly type expressions."""
    type_text = type_text.strip()
    if not type_text:
        errors.append(f"{where}: empty type expression")
        return
    # Handle simple and generic type-argument lists. The second generic argument
    # 'Argument' is resolved by the default System.Activities namespace.
    for token in re.split(r"\s*,\s*", type_text):
        token = token.strip()
        if not token:
            continue
        if token.startswith("x:"):
            name = token.split(":", 1)[1]
            if name not in X_BUILTINS:
                errors.append(f"{where}: unsupported x: CLR type '{name}'. Use a CLR namespace alias (for example sys:DateTime).")
        elif ":" in token:
            prefix, _local = token.split(":", 1)
            if prefix not in ns:
                errors.append(f"{where}: namespace prefix '{prefix}:' is not declared")
        elif token == "Argument":
            pass
        elif token in {"Object", "String", "Int32", "Boolean", "DateTime"}:
            errors.append(f"{where}: unqualified CLR type '{token}' in a generic type list")


def validate(project: Path) -> list[str]:
    errors: list[str] = []
    project = project.resolve()
    manifest_path = project / "project.json"
    if not manifest_path.exists():
        return ["project.json is missing"]
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"project.json cannot be parsed: {exc}"]

    if manifest.get("main") != "Main.xaml":
        errors.append(f"project.json main must be Main.xaml, found {manifest.get('main')!r}")
    if not (project / (manifest.get("main") or "")).is_file():
        errors.append(f"project main workflow is missing: {manifest.get('main')!r}")
    if manifest.get("targetFramework") != "Windows":
        errors.append(f"expected targetFramework Windows, found {manifest.get('targetFramework')!r}")
    if manifest.get("expressionLanguage") != "VisualBasic":
        errors.append(f"expected VisualBasic expressionLanguage, found {manifest.get('expressionLanguage')!r}")
    if manifest.get("studioVersion") != "24.10.8":
        errors.append(f"expected Studio version 24.10.8, found {manifest.get('studioVersion')!r}")
    deps = manifest.get("dependencies", {})
    for package in ("UiPath.System.Activities", "UiPath.Excel.Activities"):
        if package not in deps:
            errors.append(f"project.json dependency missing: {package}")

    xaml_paths = sorted(project.rglob("*.xaml"))
    roots: dict[str, ET.Element] = {}
    members: dict[str, dict[str, str]] = {}
    for path in xaml_paths:
        rel = path.relative_to(project).as_posix()
        raw = path.read_text(encoding="utf-8")
        try:
            root = ET.fromstring(raw)
        except Exception as exc:
            errors.append(f"{rel}: invalid XML: {exc}")
            continue
        roots[rel] = root
        ns = namespaces(path)
        if root.tag != f"{{{ACT}}}Activity":
            errors.append(f"{rel}: root must be System.Activities Activity")
        expected_class = path.relative_to(project).with_suffix("").as_posix().replace("/", "_")
        actual_class = root.attrib.get(XCLASS)
        if actual_class != expected_class:
            errors.append(f"{rel}: x:Class must be '{expected_class}', found {actual_class!r}")

        # DateTime/Guid/etc. are not built-in x: schema types.
        for match in re.finditer(r"\bx:(DateTime|DateTimeOffset|Guid)\b", raw):
            errors.append(f"{rel}: unsupported x:{match.group(1)} at offset {match.start()}; use sys:{match.group(1)}")
        if "sys:DateTime" in raw and ns.get("sys") != SYS_CLR:
            errors.append(f"{rel}: sys:DateTime is used but xmlns:sys must be '{SYS_CLR}'")

        props: dict[str, str] = {}
        for prop in root.findall(".//x:Members/x:Property", {"x": X}):
            name, signature = prop.attrib.get("Name", ""), prop.attrib.get("Type", "")
            if not name:
                errors.append(f"{rel}: x:Property has no Name")
                continue
            if name in props:
                errors.append(f"{rel}: duplicate workflow argument {name}")
            props[name] = signature
            m = re.fullmatch(r"(InArgument|OutArgument|InOutArgument)\((.+)\)", signature)
            if not m:
                errors.append(f"{rel}: argument {name} has unsupported type declaration {signature!r}")
            else:
                check_type(m.group(2), ns, f"{rel} argument {name}", errors)
        members[rel] = props

        # Validate activity prefixes and arguments to invoked child workflows.
        ids = [el.attrib[IDREF] for el in root.iter() if IDREF in el.attrib]
        seen: set[str] = set()
        for item in ids:
            if item in seen:
                errors.append(f"{rel}: duplicate WorkflowViewState.IdRef {item}")
            seen.add(item)

        # Invoke Code requires a non-empty VB block and every dictionary key must
        # be referenced by the code; verify the type wrappers are resolvable.
        for invoke_code in root.findall(".//ui:InvokeCode", {"ui": UI}):
            if invoke_code.attrib.get("Language") != "VBNet":
                errors.append(f"{rel}: InvokeCode Language must be VBNet")
            code = invoke_code.attrib.get("Code", "")
            if not code.strip():
                errors.append(f"{rel}: InvokeCode has empty Code")
            arg_parent = invoke_code.find("ui:InvokeCode.Arguments", {"ui": UI})
            if arg_parent is not None and len(arg_parent):
                for arg in list(arg_parent[0]):
                    key = arg.attrib.get(XKEY, "")
                    if key and not re.search(rf"\b{re.escape(key)}\b", code):
                        errors.append(f"{rel}: InvokeCode argument {key!r} is not referenced by Code")
                    type_arg = arg.attrib.get(f"{{{X}}}TypeArguments", "")
                    if type_arg:
                        check_type(type_arg, ns, f"{rel} InvokeCode argument {key}", errors)

        # Validate every XAML generic type argument, including Sequence variables,
        # Invoke Code arguments, and invoked-workflow argument dictionary values.
        for el in root.iter():
            type_args = el.attrib.get(f"{{{X}}}TypeArguments", "")
            if type_args:
                check_type(type_args, ns, f"{rel} {el.tag.split('}')[-1]} x:TypeArguments", errors)

        # All Workbook Read/Write activities must use UiPath's standard activity
        # namespace, not an ad-hoc CLR prefix.
        for el in root.iter():
            local = el.tag.split("}")[-1]
            if local in {"ReadRange", "WriteRange"} and el.tag.startswith("{clr-namespace:"):
                errors.append(f"{rel}: {local} must use the UiPath activity namespace (ui:{local})")
            if local in {"ReadRange", "WriteRange"} and el.tag.startswith("{http://schemas.uipath.com/workflow/activities}"):
                pass

    # Workflow invocation contract: path, argument names, directions and CLR types.
    for rel, root in roots.items():
        for invoke in root.findall(".//ui:InvokeWorkflowFile", {"ui": UI}):
            target = (invoke.attrib.get("WorkflowFileName") or "").replace("\\", "/")
            if not target:
                errors.append(f"{rel}: InvokeWorkflowFile has no WorkflowFileName")
                continue
            if target not in roots:
                errors.append(f"{rel}: workflow reference not found: {target}")
                continue
            declared = members.get(target, {})
            parent = invoke.find("ui:InvokeWorkflowFile.Arguments", {"ui": UI})
            passed: set[str] = set()
            if parent is not None and len(parent):
                for arg in list(parent[0]):
                    name = arg.attrib.get(XKEY, "")
                    passed.add(name)
                    direction = arg.tag.split("}")[-1]
                    type_arg = arg.attrib.get(f"{{{X}}}TypeArguments", "")
                    signature = f"{direction}({type_arg})"
                    if name not in declared:
                        errors.append(f"{rel}: {target} receives undeclared argument {name!r}")
                    elif declared[name] != signature:
                        errors.append(f"{rel}: {target}.{name} type/direction mismatch: caller {signature}, child {declared[name]}")
            for missing in sorted(set(declared) - passed):
                errors.append(f"{rel}: {target} argument {missing!r} is not supplied by caller")

    # Main workflow must use supported Workbook activities under the standard UI alias.
    main = roots.get("Main.xaml")
    if main is not None:
        if not main.findall(".//ui:ReadRange", {"ui": UI}):
            errors.append("Main.xaml: no ui:ReadRange activities found")
        if not main.findall(".//ui:WriteRange", {"ui": UI}):
            errors.append("Main.xaml: no ui:WriteRange activities found")

    mapping = project / "Migration" / "Function_Mapping.csv"
    if not mapping.exists():
        errors.append("Migration/Function_Mapping.csv is missing")
    else:
        try:
            with mapping.open(encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f))
            if len(rows) < 200:
                errors.append(f"Function_Mapping.csv contains only {len(rows)} rows; expected full source traceability")
            for row in rows:
                target = (row.get("Target XAML / artifact") or "").replace("\\", "/")
                if target.endswith(".xaml") and target not in roots:
                    errors.append(f"mapping references missing XAML: {target}")
        except Exception as exc:
            errors.append(f"Function_Mapping.csv cannot be read: {exc}")

    if not (project / "Skills" / "Applied_Skill_References.md").exists():
        errors.append("Skills/Applied_Skill_References.md is missing")
    if not (project / "Skills" / "Library" / "INDEX.md").exists():
        errors.append("Skills/Library/INDEX.md is missing")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("project", nargs="?", type=Path, default=Path(__file__).parent)
    args = parser.parse_args()
    errors = validate(args.project)
    if errors:
        print(f"FAILED: {len(errors)} issue(s)")
        for item in errors:
            print("-", item)
        return 1
    print("PASS: project manifest, all XAML XML, x:Class, DateTime CLR namespace, argument types, unique IdRefs, invoked workflow paths/contracts, InvokeCode bindings, Excel activity namespace, package dependencies and source traceability.")
    print("NOTE: This is a static structural check only. UiPath Studio/CLI Validate + Build is still required to verify the actual UiPath serializer, package activity metadata and VB expression compilation.")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
