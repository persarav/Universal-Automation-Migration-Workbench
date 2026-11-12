"""Safe importer/indexer for Agent Skills-style folders and GitHub skill repositories.

Imported scripts are indexed as reference files only; they are never executed.
"""
from __future__ import annotations
import io, json, re, zipfile
from pathlib import Path
from typing import Any
import requests

DEFAULT_SOURCES = [
    {"name":"Microsoft Power Automate skills", "url":"https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate", "kind":"platform skill pack", "focus":"Cloud flow creation, connector operation discovery, debugging, validation, desktop flows, environment routing", "status":"reference source"},
    {"name":"UiPath skills", "url":"https://github.com/UiPath/skills", "kind":"platform skill repository", "focus":"UiPath XAML and coded workflow authoring, project validation, packaging, testing, platform operations, agents and troubleshooting", "status":"reference source; repository marks some skills preview/in-development"},
    {"name":"VBA skills by RenRMT", "url":"https://github.com/RenRMT/claude-code-vba-skills", "kind":"language skill repository", "focus":"VBA coding conventions, types, data management and maintainability", "status":"reference source"},
    {"name":"Python skills topic", "url":"https://github.com/topics/python-skills", "kind":"community topic directory", "focus":"Community-maintained Python engineering, code quality, testing and agent skills", "status":"not a single canonical skill pack"},
    {"name":"Agent Skills specification", "url":"https://github.com/agentskills/agentskills", "kind":"skill format specification", "focus":"Portable SKILL.md folders with optional scripts, references and assets", "status":"format specification, not platform implementation skills"},
]


def _parse_frontmatter(text: str) -> dict[str, str]:
    m = re.match(r"\A---\s*\n(.*?)\n---\s*\n?(.*)\Z", text, re.S)
    if not m: return {"name":"", "description":"", "content":text}
    meta = {}
    try:
        import yaml
        loaded = yaml.safe_load(m.group(1)) or {}
        if isinstance(loaded, dict): meta = {str(k): str(v) for k, v in loaded.items() if isinstance(v, (str, int, float, bool))}
    except Exception:
        # Fallback supports simple one-line metadata if PyYAML is unavailable.
        for line in m.group(1).splitlines():
            if ":" in line:
                k, v = line.split(":", 1); meta[k.strip()] = v.strip().strip("\"'")
    return {"name":meta.get("name", ""), "description":meta.get("description", ""), "content":m.group(2)}


def parse_skill_markdown(text: str, source: str = "uploaded") -> dict[str, Any]:
    meta = _parse_frontmatter(text)
    return {"name":meta.get("name") or "Imported skill", "description":meta.get("description", ""), "content":meta.get("content", text), "source":source}


def import_skill_zip(raw: bytes, label: str = "uploaded skill archive", max_files: int = 300, max_total_bytes: int = 12_000_000) -> list[dict[str, Any]]:
    skills=[]; total=0
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        files=[x for x in zf.infolist() if not x.is_dir()]
        if len(files)>max_files: raise ValueError(f"Skill archive exceeds {max_files} files.")
        for info in files:
            p=Path(info.filename)
            if info.file_size>1_000_000: continue
            total += info.file_size
            if total>max_total_bytes: break
            # Read Markdown references only. Scripts/assets are not executed or ingested by default.
            if p.name.lower()=="skill.md" or (p.suffix.lower()==".md" and any(x in p.parts for x in ("references", "docs"))):
                text=zf.read(info).decode("utf-8", "replace")
                skill=parse_skill_markdown(text, f"{label}:{info.filename}")
                skill["name"] = skill["name"] or p.parent.name
                skills.append(skill)
    return skills


def parse_github_repo_url(url: str) -> tuple[str,str,str,str]:
    m=re.match(r"https?://github\.com/([^/]+)/([^/]+)(?:/tree/([^/]+)(/.*)?)?", url.strip())
    if not m: raise ValueError("Use a public GitHub repository URL, optionally pointing to /tree/branch/path. A GitHub topic page is a directory, not a repository.")
    owner, repo, branch, subpath=m.groups()
    return owner, repo.removesuffix(".git"), branch or "main", (subpath or "").lstrip("/")


def import_github_skill_repo(url: str, token: str = "", max_files: int = 200, max_bytes: int = 10_000_000) -> list[dict[str, Any]]:
    owner, repo, branch, prefix = parse_github_repo_url(url)
    headers={"Accept":"application/vnd.github+json", "X-GitHub-Api-Version":"2022-11-28"}
    if token: headers["Authorization"]="Bearer " + token
    api=f"https://api.github.com/repos/{owner}/{repo}/git/trees/{branch}?recursive=1"
    r=requests.get(api, headers=headers, timeout=35)
    if not r.ok: raise RuntimeError(f"GitHub tree fetch failed ({r.status_code}): {r.text[:500]}")
    tree=r.json().get("tree", [])
    md_paths=[x["path"] for x in tree if x.get("type")=="blob" and x["path"].lower().endswith(".md") and (not prefix or x["path"].startswith(prefix.rstrip("/")+"/")) and (Path(x["path"]).name.lower() in {"skill.md", "readme.md", "specification.md"} or "/references/" in "/"+x["path"].lower()+"/" or "/docs/" in "/"+x["path"].lower()+"/")]
    if len(md_paths)>max_files: md_paths=md_paths[:max_files]
    skills=[]; total=0
    for path in md_paths:
        raw_url=f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{path}"
        rr=requests.get(raw_url, headers={"Authorization": f"Bearer {token}"} if token else {}, timeout=25)
        if not rr.ok: continue
        content=rr.text
        total += len(content.encode("utf-8"))
        if total>max_bytes: break
        if Path(path).name.lower()=="skill.md":
            skill=parse_skill_markdown(content, f"github:{owner}/{repo}/{path}")
            skill["name"] = skill["name"] or Path(path).parent.name
            skills.append(skill)
        else:
            skills.append({"name":Path(path).stem, "description":"Reference document imported from skill repository", "content":content, "source":f"github:{owner}/{repo}/{path}"})
    return skills


def select_relevant_skills(skills: list[dict[str, Any]], query: str, limit: int=6) -> list[dict[str, Any]]:
    terms=set(re.findall(r"[a-z0-9_+#.-]{3,}", query.lower()))
    scored=[]
    for s in skills:
        text=(s.get("name","")+" "+s.get("description","")+" "+s.get("content","")).lower()
        score=sum(1 for t in terms if t in text)
        if score: scored.append((score,s))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [s for _,s in scored[:limit]]
