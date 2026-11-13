from __future__ import annotations
import io, json, os, re, zipfile, textwrap
from pathlib import Path
import streamlit as st
import pandas as pd

from engine import (
    TARGETS, extract_upload, analyze_source, build_intermediate_representation,
    build_migration_prompt, call_openai_compatible, generate_fallback, build_report_md,
    normalize_migration_result,
)
from skill_loader import (DEFAULT_SOURCES, import_skill_zip, import_github_skill_repo, parse_skill_markdown, select_relevant_skills, load_bundled_skills, merge_skills)

st.set_page_config(page_title="Universal Automation Migration Workbench", page_icon="🔁", layout="wide")

st.markdown("""
<style>
:root { --ink:#182536; --muted:#738196; --line:#e4eaf2; --accent:#315efb; }
.block-container {padding-top:1.35rem; padding-bottom:3rem; max-width:1440px;}
.hero {padding:1.45rem 1.6rem; border:1px solid #e4eaf2; border-radius:18px; background:linear-gradient(120deg,#f5f8ff 0%,#f8fbff 55%,#f4fbfa 100%); margin-bottom:1.2rem;}
.hero h1 {font-size:2.05rem; line-height:1.13; letter-spacing:-.04em; color:#16253a; margin:0 0 .55rem 0;}
.hero p {color:#627186; font-size:1rem; margin:0; max-width:900px;}
.eyebrow {font-size:.72rem; color:#315efb; text-transform:uppercase; font-weight:800; letter-spacing:.13em; margin-bottom:.55rem;}
.panel {border:1px solid #e4eaf2; border-radius:15px; padding:1rem 1.1rem; background:white;}
.small-muted {color:#738196; font-size:.82rem;}
.step-pill {display:inline-block; padding:.28rem .7rem; border-radius:30px; background:#edf2ff; color:#3152c9; font-weight:700; font-size:.78rem; margin-right:.35rem;}
div[data-testid="stMetric"] {background:#fff; border:1px solid #e7edf4; padding:.65rem .8rem; border-radius:13px;}
.stButton>button {border-radius:10px; font-weight:650;}
</style>
""", unsafe_allow_html=True)

if "skills" not in st.session_state: st.session_state.skills=load_bundled_skills()
if "extracted" not in st.session_state: st.session_state.extracted=None
if "analysis" not in st.session_state: st.session_state.analysis=None
if "ir" not in st.session_state: st.session_state.ir=None
if "migration" not in st.session_state: st.session_state.migration=None
if not isinstance(st.session_state.get("artifacts"), dict): st.session_state.artifacts={}
if "page" not in st.session_state: st.session_state.page="Migrate"

with st.sidebar:
    st.markdown("## 🔁 Migration Workbench")
    st.caption("Understand first. Map second. Generate third.")
    page=st.radio("Workspace", ["Migrate", "Skill Library", "Architecture & Limits"], index=["Migrate","Skill Library","Architecture & Limits"].index(st.session_state.page))
    st.session_state.page=page
    st.divider()
    st.markdown("**Engine status**")
    st.markdown("<span style='color:#16835e'>●</span> Static analyzer available", unsafe_allow_html=True)
    st.markdown(f"<span style='color:#16835e'>●</span> {len(st.session_state.skills)} bundled/imported skill Markdown files", unsafe_allow_html=True)
    st.caption("AI generation is optional. Static analysis runs locally in this app process.")

st.markdown('<div class="hero"><div class="eyebrow">Code intelligence · RPA modernization · platform migration</div><h1>Universal Automation Migration Workbench</h1><p>Reverse-engineer code and RPA workflows, extract their business rules, map them to a universal workflow model, and generate a target-platform migration package.</p></div>', unsafe_allow_html=True)

if page == "Migrate":
    st.markdown('<span class="step-pill">01 · SOURCE</span><span class="step-pill">02 · UNDERSTAND</span><span class="step-pill">03 · MIGRATE</span><span class="step-pill">04 · VALIDATE</span>', unsafe_allow_html=True)
    left, right = st.columns([1.04, .96], gap="large")
    with left:
        st.subheader("1. Source project")
        uploads=st.file_uploader("Upload source files or a project archive", type=["xlsm","xlam","bas","cls","frm","vba","txt","py","pyw","vb","vbs","cs","xaml","json","xml","flow","atmx","bprelease","uipx","nupkg","zip","md","mdx"], accept_multiple_files=True, help="Supports VBA exports and macro-enabled Excel workbooks, Python, VB.NET, C#, UiPath XAML, JSON flow definitions, XML exports and skill archives.")
        manual = st.text_area("Or paste source code / process description", height=170, placeholder="Paste code, exported workflow XML/JSON, or describe the automation when there is no source file.")
        source_hint=st.selectbox("Source technology (optional hint)", ["Auto-detect","VBA / Excel Macro","Python","VB.NET / VBScript","C# / .NET","UiPath XAML","Power Automate cloud flow","Power Automate Desktop","Automation Anywhere","Blue Prism","Mixed / other"])
        user_context=st.text_area("Business context and migration requirements (recommended)", height=100, placeholder="What starts the process? What are the inputs/outputs? Any rules that must not change? Target platform/version? Environment restrictions?")
    with right:
        st.subheader("2. Target platform")
        target=st.selectbox("Migrate into", TARGETS, index=0)
        migration_mode=st.radio("Migration mode", ["Generate offline migration (no API key)", "Understand + explain", "Optional AI semantic enhancement"], index=0, help="Offline migration uses bundled skill rules and deterministic source parsers. AI is optional and never required to generate the SYNC UiPath project.")
        with st.expander("Optional AI enhancement (not required)", expanded=False):
            env_key=os.getenv("OPENAI_API_KEY", "")
            api_key=st.text_input("Model API key", value=env_key, type="password", help="Optional. Never needed for offline analysis or the deterministic SYNC-to-UiPath migration.")
            model=st.text_input("Model name", value=os.getenv("MIGRATION_MODEL", "gpt-4.1"))
            base_url=st.text_input("OpenAI-compatible API base URL", value=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"))
            send_to_ai=st.checkbox("I explicitly approve sending source code and selected skills to the model provider", value=False)
        if "api_key" not in locals(): api_key=env_key if "env_key" in locals() else os.getenv("OPENAI_API_KEY", "")
        if "send_to_ai" not in locals(): send_to_ai=False
        st.caption("Offline analysis and SYNC UiPath project generation run without any model API key.")
        chosen_skill_limit=st.slider("Relevant imported skill references included in this migration", 1, 30, 12)
        if st.session_state.skills:
            st.caption(f"Skill library contains {len(st.session_state.skills)} bundled/imported Markdown references, available without an API key.")
        else:
            st.caption("Bundled platform playbooks are loaded. Import the full upstream Markdown repositories from Skill Library for broader coverage.")

    c1,c2=st.columns([1,4])
    with c1:
        analyze_clicked=st.button("Analyze source", type="primary", use_container_width=True)
    with c2:
        st.caption("The analyzer never executes uploaded macros, scripts, bots or workflows.")

    if analyze_clicked:
        files=[]; notes=[]
        try:
            for up in uploads or []:
                item=extract_upload(up.name, up.getvalue())
                files.extend(item["files"]); notes.extend(item.get("notes", []))
            if manual.strip():
                ext={"VBA / Excel Macro":".bas","Python":".py","VB.NET / VBScript":".vb","C# / .NET":".cs","UiPath XAML":".xaml","Power Automate cloud flow":".json"}.get(source_hint,".txt")
                files.append({"filename":"pasted_source"+ext,"text":manual,"language":source_hint if source_hint!="Auto-detect" else "Unknown / mixed source"})
            if not files:
                st.error("Upload at least one source file or paste code before analysis.")
            else:
                extracted={"filename":", ".join([x.name for x in uploads or []]) or "Pasted source", "files":files,"notes":notes}
                analysis=analyze_source(extracted)
                ir=build_intermediate_representation(analysis)
                st.session_state.extracted=extracted; st.session_state.analysis=analysis; st.session_state.ir=ir; st.session_state.migration=None; st.session_state.artifacts={}
                st.success(f"Source analyzed: {len(files)} file(s), {analysis['symbol_count']} procedure/function/class symbol(s) detected.")
                for n in notes: st.info(n)
        except Exception as exc:
            st.exception(exc)

    if st.session_state.analysis:
        analysis=st.session_state.analysis; ir=st.session_state.ir
        st.divider(); st.subheader("Source understanding")
        a,b,c,d=st.columns(4)
        a.metric("Files analyzed", len(analysis.get("files",[])))
        b.metric("Procedures / functions", analysis.get("symbol_count",0))
        c.metric("Modules / files", len(analysis.get("modules",[])))
        d.metric("External systems", len(analysis.get("external_touchpoints",[])))
        if analysis.get("files"):
            st.dataframe(pd.DataFrame(analysis["files"]), use_container_width=True, hide_index=True)
        tabs=st.tabs(["Procedure inventory","Logic indicators","External touchpoints","Workbook / workflow details","Intermediate model"])
        with tabs[0]:
            syms=analysis.get("symbols",[])
            if syms:
                df=pd.DataFrame([{ "Name":x["name"],"Kind":x["kind"],"File":x["source_file"],"Line":x["line"],"Signature":x["signature"],"Calls":", ".join(x.get("calls") or [])} for x in syms])
                st.dataframe(df, use_container_width=True, hide_index=True, height=360)
                with st.expander("View extracted code excerpts"):
                    names=[f"{x['name']} — {x['source_file']}:{x['line']}" for x in syms]
                    selected=st.selectbox("Procedure", names)
                    idx=names.index(selected); st.code(syms[idx].get("body_excerpt", ""), language="vb")
            else: st.info("No named procedures were found in the uploaded source. Use XML/JSON inspection or provide a source export.")
        with tabs[1]:
            if analysis.get("patterns"):
                st.dataframe(pd.DataFrame([{"Pattern":k,"Occurrences":v} for k,v in analysis["patterns"].items()]).sort_values("Occurrences",ascending=False), use_container_width=True, hide_index=True)
            else: st.info("No structural indicators detected.")
            st.caption("These counts are heuristic indicators, not proof of the full runtime control flow.")
        with tabs[2]:
            st.write(analysis.get("external_touchpoints") or ["No known external touchpoints detected by keyword scan."])
            if analysis.get("secret_findings"): st.warning("Potential credentials or tokens detected. Remove/rotate secrets before using any external model."); st.json(analysis["secret_findings"])
            else: st.success("No obvious hard-coded credential patterns detected. This is not a complete secret scan.")
        with tabs[3]:
            if analysis.get("notes"): st.write(analysis["notes"])
            if analysis.get("xaml"): st.json(analysis["xaml"][:2])
            if analysis.get("json_workflows"): st.json(analysis["json_workflows"][:2])
            if not (analysis.get("notes") or analysis.get("xaml") or analysis.get("json_workflows")): st.info("No workbook metadata or native workflow metadata to display.")
        with tabs[4]:
            st.caption("The universal intermediate representation is an initial evidence model. Semantic details need confirmation or AI analysis.")
            st.json(ir)

        st.divider(); st.subheader("3. Generate migration package")
        if st.button("Generate migration", type="primary", use_container_width=True):
            skill_query = f"{target} {user_context} {analysis.get('source_filename','')} {' '.join(x.get('name', '') for x in analysis.get('symbols', [])[:100])} {' '.join(analysis.get('external_touchpoints', []))}"
            relevant=select_relevant_skills(st.session_state.skills, skill_query, limit=chosen_skill_limit)
            if not relevant:
                relevant=st.session_state.skills[:chosen_skill_limit]
            try:
                if migration_mode == "Understand + explain":
                    result=generate_fallback("Platform-neutral specification", ir, analysis)
                    result=normalize_migration_result(result, "Platform-neutral specification", analysis)
                    result["generation_mode"]="Offline source understanding; no API key"
                else:
                    # Always create the deterministic target package first. For the
                    # SYNC workbook this executes a source-specific translator that
                    # emits the modular UiPath XAML project without an LLM.
                    result=generate_fallback(target, ir, analysis, relevant_skills=relevant)
                    result=normalize_migration_result(result, target, analysis)
                    result["generation_mode"]="Offline deterministic generation; no API key"
                    if migration_mode == "Optional AI semantic enhancement" and send_to_ai and api_key.strip():
                        try:
                            prompt=build_migration_prompt(analysis, ir, target, relevant, user_context)
                            with st.spinner("Adding an optional semantic review; deterministic target files are preserved…"):
                                ai_result=call_openai_compatible(prompt, st.session_state.extracted["files"], api_key.strip(), model.strip(), base_url.strip())
                            ai_result=normalize_migration_result(ai_result, target, analysis)
                            review=["# Optional AI semantic review", "", str(ai_result.get("executive_summary", "")), "", "## Additional unresolved items"]
                            review.extend("- " + str(x) for x in ai_result.get("unresolved_items", []))
                            review.extend(["", "## Additional test recommendations"])
                            review.extend("- " + str(x) for x in ai_result.get("test_plan", []))
                            result.setdefault("generated_files", []).append({"path":"Migration/AI_Semantic_Review.md", "content":"\n".join(review), "purpose":"Optional model-generated review; it does not replace deterministic XAML."})
                            result["generation_mode"]="Offline deterministic project + optional AI review"
                        except Exception as ai_exc:
                            st.warning(f"Optional AI review was unavailable; the offline migration project is complete: {ai_exc}")
                    elif migration_mode == "Optional AI semantic enhancement" and not (send_to_ai and api_key.strip()):
                        st.info("No API key or consent supplied. Generated the complete offline migration instead.")
                result=normalize_migration_result(result, target, analysis)
                # Normalize all generator output before any metadata or rendering use.
                skill_md = ["# Skill references used by this migration", "", "These bundled/imported Markdown files informed the deterministic migration design. They are reference guidance, not executable code.", ""]
                for skill in relevant:
                    skill_md.extend([f"## {skill.get('name','Skill')}", f"Source: {skill.get('source','unknown')}", "", str(skill.get('description','')), ""])
                result.setdefault("generated_files", []).append({"path":"Skill_References/Applied_Skills.md", "content":"\n".join(skill_md), "purpose":"Traceability for the skill guidance selected during this migration."})
                # Preserve every bundled/imported Markdown skill inside the exported
                # migration package. The relevant subset is separately embedded in
                # Skills/Applied_Skill_References.md by the target generator.
                skill_index=["# Complete Skill Library Snapshot", "", f"Skill Markdown files included: {len(st.session_state.skills)}", "", "All imported Markdown is stored as reference data; scripts are not executed.", ""]
                for i, skill in enumerate(st.session_state.skills):
                    base=re.sub(r"[^A-Za-z0-9_.-]+", "_", str(skill.get("name") or f"skill_{i}"))[:80] or f"skill_{i}"
                    digest=re.sub(r"[^A-Za-z0-9_.-]+", "_", str(skill.get("source") or "bundled"))[-55:] or "local"
                    rel=f"Skills/Library/{i+1:03d}_{base}_{digest}.md"
                    result["generated_files"].append({"path":rel, "content":str(skill.get("content") or ""), "purpose":"Offline Markdown skill/reference included in the project library."})
                    skill_index.append(f"- `{rel}` — {skill.get('description','')}")
                result["generated_files"].append({"path":"Skills/Library/INDEX.md", "content":"\n".join(skill_index), "purpose":"Index of every Markdown skill packaged with the migration."})
                report=build_report_md(result, analysis, target)
                artifacts={"migration-report.md":report,"source-analysis.json":json.dumps(analysis,indent=2,ensure_ascii=False),"workflow-ir.json":json.dumps(ir,indent=2,ensure_ascii=False)}
                for f in result.get("generated_files",[]):
                    path=Path(f.get("path","generated.txt"))
                    if path.is_absolute() or ".." in path.parts: continue
                    artifacts[path.as_posix()]=str(f.get("content",""))
                st.session_state.migration=result; st.session_state.artifacts=artifacts; st.session_state.target=target
                if target == "UiPath (XAML)" and str((result.get("target_design") or {}).get("status") or "").startswith("modular business-logic"):
                    st.success(f"Offline UiPath project generated: {len([x for x in result['generated_files'] if x['path'].lower().endswith('.xaml')])} XAML workflows; {len(result['mapping_table'])} source procedures mapped. No API key used.")
                else:
                    st.success("Offline migration package generated. Review the mapping and target-specific requirements before use.")
            except Exception as exc:
                st.error(f"Migration generation failed: {type(exc).__name__}: {exc}")
        if st.session_state.migration:
            result=normalize_migration_result(st.session_state.get("migration"), st.session_state.get("target", target), analysis)
            st.session_state.migration=result
            artifact_state=st.session_state.get("artifacts")
            if not isinstance(artifact_state, dict):
                artifact_state={}
                st.session_state.artifacts=artifact_state
            st.markdown(f"### {result.get('executive_summary','Migration output')}")
            st.caption(f"Generation mode: {result.get('generation_mode','unknown')} · Confidence: {result.get('confidence','not rated')}")
            mtabs=st.tabs(["Mapping","Unresolved items","Test plan","Risks & assumptions","Generated files"])
            with mtabs[0]:
                # Keep the rendering boundary defensive: legacy session data, a failed
                # generator, or a provider response must never break the whole page.
                mapping_rows=result.get("mapping_table") if isinstance(result, dict) and isinstance(result.get("mapping_table"), list) else []
                mapping_rows=[row for row in mapping_rows if isinstance(row, dict)]
                if mapping_rows: st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)
                else: st.info("No mapping table was produced by this analyzer. See the procedure inventory and unresolved items.")
            with mtabs[1]: st.write(result.get("unresolved_items",[]))
            with mtabs[2]: st.write(result.get("test_plan",[]))
            with mtabs[3]:
                st.markdown("**Risks**"); st.write(result.get("risks",[])); st.markdown("**Assumptions**"); st.write(result.get("assumptions",[]))
            with mtabs[4]:
                for path, content in artifact_state.items():
                    if not isinstance(path, str) or not isinstance(content, str):
                        continue
                    with st.expander(path):
                        st.code(content[:20000], language="xml" if path.endswith(".xaml") else "python" if path.endswith(".py") else "json" if path.endswith(".json") else "markdown")
            zip_buffer=io.BytesIO()
            with zipfile.ZipFile(zip_buffer,"w",zipfile.ZIP_DEFLATED) as z:
                for path, content in artifact_state.items():
                    if isinstance(path, str) and isinstance(content, str): z.writestr(path,content)
                z.writestr("README_FIRST.md", "Review migration-report.md first. For the SYNC workbook, the XAML package contains deterministic offline business rules in modular workflows. Open project.json in UiPath Studio, restore packages, validate and build. SAP GUI selectors must be captured in the approved SAP environment before live SAP pulling can run. Regression-test on a copy of the workbook before enabling write-back.\n")
            st.download_button("Download migration package (.zip)", data=zip_buffer.getvalue(), file_name="automation_migration_package.zip", mime="application/zip", type="primary", use_container_width=True)
            json_buf=json.dumps(result,indent=2,ensure_ascii=False).encode("utf-8")
            st.download_button("Download migration result (.json)", data=json_buf, file_name="migration_result.json", mime="application/json")

elif page == "Skill Library":
    st.subheader("Skill library")
    st.write("The workbench bundles local platform playbooks and can import every Markdown file from the linked upstream repositories. Imported scripts and binaries are never executed.")
    st.markdown("#### Supplied skill sources")
    st.dataframe(pd.DataFrame(DEFAULT_SOURCES), use_container_width=True, hide_index=True)
    st.info(f"{len(st.session_state.skills)} Markdown skill/reference files are currently loaded, including bundled offline playbooks. No model API key is needed to use them.")
    st.caption("GitHub topic pages are discovery indexes, not a single skill repository. Select/import individual Python or agentic skill repositories from those indexes. The importer stores Markdown only and does not execute imported scripts.")
    if st.button("Fetch all Markdown skills from supplied GitHub repositories (no API key)"):
        repo_sources = [
            ("UiPath RPA", "https://github.com/UiPath/skills/tree/main/skills"),
            ("Microsoft Power Automate", "https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate"),
            ("VBA", "https://github.com/RenRMT/claude-code-vba-skills"),
            ("Agent Skills specification", "https://github.com/agentskills/agentskills"),
        ]
        added=0; errors=[]
        progress=st.progress(0.0, text="Starting skill repository imports…")
        for i,(label,url) in enumerate(repo_sources, start=1):
            try:
                imported=import_github_skill_repo(url, max_files=1200, max_bytes=45_000_000)
                st.session_state.skills=merge_skills(st.session_state.skills, imported)
                added += len(imported)
                progress.progress(i/len(repo_sources), text=f"Imported {label}: {len(imported)} Markdown files")
            except Exception as exc:
                errors.append(f"{label}: {type(exc).__name__}: {exc}")
                progress.progress(i/len(repo_sources), text=f"Could not import {label}")
        if added: st.success(f"Imported {added} Markdown skill/reference files without an API key.")
        for err in errors: st.warning(err)
        if not added and errors: st.info("Bundled local skills remain available. Retry the import later if GitHub rate limits or network access prevented downloads.")
    st.markdown("#### Import a GitHub skill repository")
    repo_url=st.text_input("Public GitHub repository or /tree/branch/path URL", placeholder="https://github.com/owner/repo")
    gh_token=st.text_input("Optional GitHub token (private repo / higher API limits)", type="password")
    if st.button("Fetch skill pack from GitHub"):
        try:
            with st.spinner("Reading skill manifests and Markdown references…"):
                imported=import_github_skill_repo(repo_url,gh_token)
            st.session_state.skills=merge_skills(st.session_state.skills, imported)
            st.success(f"Imported {len(imported)} skill/reference document(s).")
        except Exception as exc: st.error(str(exc))
    st.markdown("#### Import skill files")
    skill_zip=st.file_uploader("Skill repository ZIP", type=["zip"], key="skillzip")
    skill_md=st.file_uploader("Or upload individual SKILL.md / reference Markdown", type=["md"], accept_multiple_files=True, key="skillmd")
    if st.button("Add uploaded skill files"):
        imported=[]
        try:
            if skill_zip: imported.extend(import_skill_zip(skill_zip.getvalue(), skill_zip.name))
            for f in skill_md or []: imported.append(parse_skill_markdown(f.getvalue().decode("utf-8","replace"), f.name))
            st.session_state.skills=merge_skills(st.session_state.skills, imported)
            st.success(f"Added {len(imported)} skill/reference document(s).")
        except Exception as exc: st.error(str(exc))
    st.markdown("#### Imported library")
    if st.session_state.skills:
        st.dataframe(pd.DataFrame([{ "Name":x.get("name"),"Description":x.get("description"),"Source":x.get("source")} for x in st.session_state.skills]), use_container_width=True, hide_index=True)
        names=[f"{x.get('name')} — {x.get('source')}" for x in st.session_state.skills]
        selected=st.selectbox("Preview reference",names)
        st.markdown(st.session_state.skills[names.index(selected)].get("content","")[:12000])
        if st.button("Clear imported skills"): st.session_state.skills=load_bundled_skills(); st.rerun()
        skill_zip_buffer=io.BytesIO()
        with zipfile.ZipFile(skill_zip_buffer, "w", zipfile.ZIP_DEFLATED) as skill_archive:
            for index, skill in enumerate(st.session_state.skills):
                safe_name=re.sub(r"[^A-Za-z0-9_.-]+", "_", str(skill.get("name") or f"skill_{index}"))
                source_name=str(skill.get("source") or "bundled")
                source_slug=re.sub(r"[^A-Za-z0-9_.-]+", "_", source_name)[:100]
                skill_archive.writestr(f"{safe_name}/{source_slug}.md", str(skill.get("content", "")))
        st.download_button("Download all loaded skill Markdown (.zip)", data=skill_zip_buffer.getvalue(), file_name="Universal_Migration_Skill_Library.zip", mime="application/zip")
    else: st.info("No skills loaded. The bundled skillpacks directory should supply the offline playbooks.")
    st.warning("A GitHub 'topic' page such as python-skills is a discovery directory, not one canonical skill pack. Select and import specific repositories from it. Check each repository's license and quality before using it in production.")

else:
    st.subheader("Architecture and current boundaries")
    st.markdown("""
### What this workbench does
- Parses common source types into evidence: files, symbols, XAML activities, JSON actions, conditional/loop/error indicators and likely external system touchpoints.
- Builds a platform-neutral workflow representation before target generation.
- Bundles offline `.md` playbooks for UiPath, Power Automate, VBA, Python, Automation Anywhere, Blue Prism and Agent Skills. The Skill Library can fetch all Markdown files under the specified public repository paths without an AI key; fetched documents stay in the current Streamlit session until downloaded/exported. Imported scripts are never run.
- Generates a migration report, source-to-target mapping, unresolved-item list, test plan and target-specific packages. For the SYNC workbook, the offline translator emits one Main XAML plus reusable business-capability workflows. Native Workbook Read/Write and Invoke Workflow activities orchestrate each stage; dense DataTable rules are in bounded VB.NET Invoke Code blocks. It does not create one XAML for every tiny VBA helper.
- Uses an opt-in OpenAI-compatible model call for deeper semantic interpretation and generation.

### What it does not promise
- There is no universal one-click compiler from every language/RPA product to every other platform. Proprietary packages, selectors, credentials, connector schemas, versioned activity packages and tenant resources require target-specific information.
- A `.xlsm` workbook can contain formulas, named ranges, forms, buttons, ActiveX controls, external links, workbook events and macro code. This version uses a bundled pure-Python VBA project extractor and openpyxl metadata reader; no `oletools` or model API key is required for supported workbooks. It does not emulate Excel or execute a macro.
- Automation Anywhere `.atmx`, Blue Prism `.bprelease`, proprietary UiPath package/project archives and some Power Automate exports may require version-specific parsers. Unknown files are treated as source evidence, not fabricated as native projects.
- Static XAML/XML validation is not a UiPath compile. A generated project must be opened, packages restored, validated, built and regression-tested in the actual target product/version before deployment.

### Recommended migration quality gates
1. Source inventory and dependency discovery.
2. Business-rule reconstruction with source references.
3. Intermediate workflow model review by a process owner.
4. Target capability and connector mapping.
5. Generation with unresolved placeholders explicitly marked.
6. Static validation and target-native build.
7. Regression testing against baseline source outputs and side effects.
8. Security review, credential binding, deployment approval and rollback plan.

### Skill-pack architecture
Agent Skills are directory-based packages with a required `SKILL.md` and optional scripts, references and assets. This app imports skill instructions and Markdown references as data, uses relevance selection to include a limited amount in a migration prompt, and avoids executing third-party scripts. Skill source repositories may change; refresh imported packs when the upstream project changes.
""")

st.divider()
st.caption("Local-first migration environment · Offline deterministic SYNC-to-UiPath conversion · Optional model-assisted semantic review · Studio validation required")
