from __future__ import annotations
import io, json, os, zipfile, textwrap
from pathlib import Path
import streamlit as st
import pandas as pd

from engine import (
    TARGETS, extract_upload, analyze_source, build_intermediate_representation,
    build_migration_prompt, call_openai_compatible, generate_fallback, build_report_md,
)
from skill_loader import DEFAULT_SOURCES, import_skill_zip, import_github_skill_repo, parse_skill_markdown, select_relevant_skills

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

if "skills" not in st.session_state: st.session_state.skills=[]
if "extracted" not in st.session_state: st.session_state.extracted=None
if "analysis" not in st.session_state: st.session_state.analysis=None
if "ir" not in st.session_state: st.session_state.ir=None
if "migration" not in st.session_state: st.session_state.migration=None
if "artifacts" not in st.session_state: st.session_state.artifacts={}
if "page" not in st.session_state: st.session_state.page="Migrate"

with st.sidebar:
    st.markdown("## 🔁 Migration Workbench")
    st.caption("Understand first. Map second. Generate third.")
    page=st.radio("Workspace", ["Migrate", "Skill Library", "Architecture & Limits"], index=["Migrate","Skill Library","Architecture & Limits"].index(st.session_state.page))
    st.session_state.page=page
    st.divider()
    st.markdown("**Engine status**")
    st.markdown("<span style='color:#16835e'>●</span> Static analyzer available", unsafe_allow_html=True)
    st.markdown(f"<span style='color:#16835e'>●</span> {len(st.session_state.skills)} imported skill/reference files", unsafe_allow_html=True)
    st.caption("AI generation is optional. Static analysis runs locally in this app process.")

st.markdown('<div class="hero"><div class="eyebrow">Code intelligence · RPA modernization · platform migration</div><h1>Universal Automation Migration Workbench</h1><p>Reverse-engineer code and RPA workflows, extract their business rules, map them to a universal workflow model, and generate a target-platform migration package.</p></div>', unsafe_allow_html=True)

if page == "Migrate":
    st.markdown('<span class="step-pill">01 · SOURCE</span><span class="step-pill">02 · UNDERSTAND</span><span class="step-pill">03 · MIGRATE</span><span class="step-pill">04 · VALIDATE</span>', unsafe_allow_html=True)
    left, right = st.columns([1.04, .96], gap="large")
    with left:
        st.subheader("1. Source project")
        uploads=st.file_uploader("Upload source files or a project archive", type=["xlsm","xlam","bas","cls","frm","vba","txt","py","pyw","vb","vbs","cs","xaml","json","xml","flow","atmx","bprelease","uipx","nupkg","zip","md"], accept_multiple_files=True, help="Supports VBA exports and macro-enabled Excel workbooks, Python, VB.NET, C#, UiPath XAML, JSON flow definitions, XML exports and skill archives.")
        manual = st.text_area("Or paste source code / process description", height=170, placeholder="Paste code, exported workflow XML/JSON, or describe the automation when there is no source file.")
        source_hint=st.selectbox("Source technology (optional hint)", ["Auto-detect","VBA / Excel Macro","Python","VB.NET / VBScript","C# / .NET","UiPath XAML","Power Automate cloud flow","Power Automate Desktop","Automation Anywhere","Blue Prism","Mixed / other"])
        user_context=st.text_area("Business context and migration requirements (recommended)", height=100, placeholder="What starts the process? What are the inputs/outputs? Any rules that must not change? Target platform/version? Environment restrictions?")
    with right:
        st.subheader("2. Target platform")
        target=st.selectbox("Migrate into", TARGETS, index=0)
        migration_mode=st.radio("Migration depth", ["Understand + explain", "Generate target scaffold", "Generate full AI-assisted migration draft"], index=2, help="Full AI-assisted generation requires explicit opt-in and a configured OpenAI-compatible API. A static scaffold is available without AI.")
        st.markdown("#### AI-assisted analysis (optional)")
        env_key=os.getenv("OPENAI_API_KEY", "")
        api_key=st.text_input("Model API key", value=env_key, type="password", help="Can also be supplied through OPENAI_API_KEY. Kept in this Streamlit session only; not written into project files.")
        model=st.text_input("Model name", value=os.getenv("MIGRATION_MODEL", "gpt-4.1"))
        base_url=st.text_input("OpenAI-compatible API base URL", value=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"))
        send_to_ai=st.checkbox("I explicitly approve sending the uploaded source and selected skill references to this model provider", value=False)
        st.caption("Without this checkbox, source parsing and fallback scaffolding remain local. Check your organization's policy before transmitting confidential code or workbook data.")
        chosen_skill_limit=st.slider("Relevant imported skill references", 1, 10, 5)
        if st.session_state.skills:
            st.caption(f"Skill library has {len(st.session_state.skills)} imported SKILL.md / reference documents.")
        else:
            st.caption("No external skill packs imported yet. Add them from Skill Library.")

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
            relevant=select_relevant_skills(st.session_state.skills, f"{target} {user_context} {analysis.get('source_filename','')} {' '.join(x.get('name', '') for x in analysis.get('symbols', [])[:80])} {' '.join(analysis.get('external_touchpoints', []))}", limit=chosen_skill_limit)
            try:
                if migration_mode=="Generate full AI-assisted migration draft" and send_to_ai and api_key.strip():
                    prompt=build_migration_prompt(analysis, ir, target, relevant, user_context)
                    with st.spinner("Generating target migration draft…"):
                        result=call_openai_compatible(prompt, st.session_state.extracted["files"], api_key.strip(), model.strip(), base_url.strip())
                    result["generation_mode"]="AI-assisted; provider call explicitly approved"
                elif migration_mode=="Understand + explain":
                    result=generate_fallback("Platform-neutral specification", ir, analysis)
                    result["generation_mode"]="Static analysis only"
                else:
                    result=generate_fallback(target, ir, analysis)
                    result["generation_mode"]="Static scaffold; no AI provider call"
                    if migration_mode=="Generate full AI-assisted migration draft":
                        st.warning("AI generation was not performed. Enter an API key and explicitly approve source sharing to enable semantic model-assisted conversion. A static scaffold was generated instead.")
                report=build_report_md(result, analysis, target)
                artifacts={"migration-report.md":report,"source-analysis.json":json.dumps(analysis,indent=2,ensure_ascii=False),"workflow-ir.json":json.dumps(ir,indent=2,ensure_ascii=False)}
                for f in result.get("generated_files",[]):
                    path=Path(f.get("path","generated.txt"))
                    if path.is_absolute() or ".." in path.parts: continue
                    artifacts[str(path)]=str(f.get("content",""))
                st.session_state.migration=result; st.session_state.artifacts=artifacts; st.session_state.target=target
                st.success("Migration draft generated. Review every TODO, unresolved item and platform-specific dependency before use.")
            except Exception as exc:
                st.error(f"Migration generation failed: {exc}")
        if st.session_state.migration:
            result=st.session_state.migration
            st.markdown(f"### {result.get('executive_summary','Migration output')}")
            st.caption(f"Generation mode: {result.get('generation_mode','unknown')} · Confidence: {result.get('confidence','not rated')}")
            mtabs=st.tabs(["Mapping","Unresolved items","Test plan","Risks & assumptions","Generated files"])
            with mtabs[0]: st.dataframe(pd.DataFrame(result.get("mapping_table",[])), use_container_width=True, hide_index=True) if result.get("mapping_table") else st.info("No mapping table was produced.")
            with mtabs[1]: st.write(result.get("unresolved_items",[]))
            with mtabs[2]: st.write(result.get("test_plan",[]))
            with mtabs[3]:
                st.markdown("**Risks**"); st.write(result.get("risks",[])); st.markdown("**Assumptions**"); st.write(result.get("assumptions",[]))
            with mtabs[4]:
                for path, content in st.session_state.artifacts.items():
                    with st.expander(path):
                        st.code(content[:20000], language="xml" if path.endswith(".xaml") else "python" if path.endswith(".py") else "json" if path.endswith(".json") else "markdown")
            zip_buffer=io.BytesIO()
            with zipfile.ZipFile(zip_buffer,"w",zipfile.ZIP_DEFLATED) as z:
                for path, content in st.session_state.artifacts.items(): z.writestr(path,content)
                z.writestr("README_FIRST.md", "Review migration-report.md first. Generated target files may be scaffolds and are not guaranteed import-ready. Build and test in the exact target product/version before deployment.\n")
            st.download_button("Download migration package (.zip)", data=zip_buffer.getvalue(), file_name="automation_migration_package.zip", mime="application/zip", type="primary", use_container_width=True)
            json_buf=json.dumps(result,indent=2,ensure_ascii=False).encode("utf-8")
            st.download_button("Download migration result (.json)", data=json_buf, file_name="migration_result.json", mime="application/json")

elif page == "Skill Library":
    st.subheader("Skill library")
    st.write("Import Agent Skills-compatible packs from GitHub or a ZIP file. The workbench indexes SKILL.md and Markdown references; imported scripts are not executed.")
    st.markdown("#### Supplied skill sources")
    st.dataframe(pd.DataFrame(DEFAULT_SOURCES), use_container_width=True, hide_index=True)
    st.markdown("#### Import a GitHub skill repository")
    repo_url=st.text_input("Public GitHub repository or /tree/branch/path URL", placeholder="https://github.com/owner/repo")
    gh_token=st.text_input("Optional GitHub token (private repo / higher API limits)", type="password")
    if st.button("Fetch skill pack from GitHub"):
        try:
            with st.spinner("Reading skill manifests and Markdown references…"):
                imported=import_github_skill_repo(repo_url,gh_token)
            st.session_state.skills.extend(imported)
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
            st.session_state.skills.extend(imported)
            st.success(f"Added {len(imported)} skill/reference document(s).")
        except Exception as exc: st.error(str(exc))
    st.markdown("#### Imported library")
    if st.session_state.skills:
        st.dataframe(pd.DataFrame([{ "Name":x.get("name"),"Description":x.get("description"),"Source":x.get("source")} for x in st.session_state.skills]), use_container_width=True, hide_index=True)
        names=[f"{x.get('name')} — {x.get('source')}" for x in st.session_state.skills]
        selected=st.selectbox("Preview reference",names)
        st.markdown(st.session_state.skills[names.index(selected)].get("content","")[:12000])
        if st.button("Clear imported skills"): st.session_state.skills=[]; st.rerun()
    else: st.info("No skill packs imported. The supplied URLs are catalogued above but not assumed to be fully downloaded.")
    st.warning("A GitHub 'topic' page such as python-skills is a discovery directory, not one canonical skill pack. Select and import specific repositories from it. Check each repository's license and quality before using it in production.")

else:
    st.subheader("Architecture and current boundaries")
    st.markdown("""
### What this workbench does
- Parses common source types into evidence: files, symbols, XAML activities, JSON actions, conditional/loop/error indicators and likely external system touchpoints.
- Builds a platform-neutral workflow representation before target generation.
- Loads Agent Skills-style `SKILL.md` material from public GitHub repositories or ZIPs. Markdown references can inform migration prompts; imported scripts are never run.
- Generates a migration report, source-to-target mapping, unresolved-item list, test plan and target starter artifacts.
- Uses an opt-in OpenAI-compatible model call for deeper semantic interpretation and generation.

### What it does not promise
- There is no universal one-click compiler from every language/RPA product to every other platform. Proprietary packages, selectors, credentials, connector schemas, versioned activity packages and tenant resources require target-specific information.
- A `.xlsm` workbook can contain formulas, named ranges, forms, buttons, ActiveX controls, external links, workbook events and macro code. This first version extracts VBA modules (requires `oletools`) and basic workbook metadata; it does not emulate Excel or execute a macro.
- Automation Anywhere `.atmx`, Blue Prism `.bprelease`, proprietary UiPath package/project archives and some Power Automate exports may require version-specific parsers. Unknown files are treated as source evidence, not fabricated as native projects.
- A generated XAML or flow design must be opened, validated, built and tested in the actual target product/version before deployment.

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
st.caption("Prototype migration environment · Local-first static analysis · Optional model-assisted generation · Human validation required")
