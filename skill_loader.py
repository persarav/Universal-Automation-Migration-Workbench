"""Safe importer/indexer for Agent Skills-style folders and GitHub skill repositories.

Imported scripts are indexed as reference files only; they are never executed.
"""
from __future__ import annotations
import io, json, re, zipfile
from pathlib import Path
from urllib.parse import quote
from typing import Any
import requests

DEFAULT_SOURCES = [
    {"name":"Microsoft Power Automate skills", "url":"https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate", "kind":"platform skill pack", "focus":"Cloud flow creation, connector operation discovery, debugging, validation, desktop flows, environment routing", "status":"reference source"},
    {"name":"UiPath skills", "url":"https://github.com/UiPath/skills/tree/main/skills", "kind":"platform skill repository", "focus":"UiPath XAML and coded workflow authoring, project validation, packaging, testing, platform operations, agents and troubleshooting", "status":"reference source; repository marks some skills preview/in-development"},
    {"name":"VBA skills by RenRMT", "url":"https://github.com/RenRMT/claude-code-vba-skills", "kind":"language skill repository", "focus":"VBA coding conventions, types, data management and maintainability", "status":"reference source"},
    {"name":"Python skills topic", "url":"https://github.com/topics/python-skills", "kind":"community topic directory", "focus":"Community-maintained Python engineering, code quality, testing and agent skills", "status":"not a single canonical skill pack; select individual repositories to import"},
    {"name":"Agentic skills topic", "url":"https://github.com/topics/agentic-skills", "kind":"community topic directory", "focus":"Community-maintained agentic workflow and tool-use skill repositories", "status":"not a single canonical skill pack; select individual repositories to import"},
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


def import_skill_zip(raw: bytes, label: str = "uploaded skill archive", max_files: int = 500, max_total_bytes: int = 20_000_000) -> list[dict[str, Any]]:
    """Import every Markdown file in a skill archive as reference data.

    Paths are checked for traversal; non-Markdown payloads are deliberately not
    loaded or executed. This allows all SKILL.md and Markdown reference files to
    be indexed without trusting bundled scripts or binaries.
    """
    skills=[]; total=0
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        files=[x for x in zf.infolist() if not x.is_dir()]
        if len(files)>max_files: raise ValueError(f"Skill archive exceeds {max_files} files.")
        md_files=[x for x in files if Path(x.filename).suffix.lower() in {".md", ".mdx"}]
        for info in md_files:
            p=Path(info.filename)
            if p.is_absolute() or ".." in p.parts:
                continue
            if info.file_size>1_500_000:
                continue
            total += info.file_size
            if total>max_total_bytes:
                break
            text=zf.read(info).decode("utf-8", "replace")
            skill=parse_skill_markdown(text, f"{label}:{info.filename}")
            if not skill.get("name") or skill["name"]=="Imported skill":
                skill["name"] = p.stem if p.stem.lower()!="skill" else p.parent.name
            skills.append(skill)
    return skills


def parse_github_repo_url(url: str) -> tuple[str,str,str,str]:
    m=re.match(r"https?://github\.com/([^/]+)/([^/]+)(?:/tree/([^/]+)(/.*)?)?", url.strip())
    if not m: raise ValueError("Use a public GitHub repository URL, optionally pointing to /tree/branch/path. A GitHub topic page is a directory, not a repository.")
    owner, repo, branch, subpath=m.groups()
    return owner, repo.removesuffix(".git"), branch or "main", (subpath or "").lstrip("/")


def import_github_skill_repo(url: str, token: str = "", max_files: int = 1200, max_bytes: int = 45_000_000) -> list[dict[str, Any]]:
    """Fetch all Markdown files in a public skill repository/subtree.

    Reads GitHub's recursive tree once and then downloads each Markdown blob.
    Only Markdown is ingested: imported scripts are never executed. Unauthenticated
    public access works without a GitHub token, subject to GitHub rate limits.
    """
    owner, repo, branch, prefix = parse_github_repo_url(url)
    headers={"Accept":"application/vnd.github+json", "X-GitHub-Api-Version":"2022-11-28"}
    if token: headers["Authorization"]="Bearer " + token
    api=f"https://api.github.com/repos/{owner}/{repo}/git/trees/{quote(branch, safe='')}?recursive=1"
    r=requests.get(api, headers=headers, timeout=45)
    if r.status_code == 404 and branch == "main":
        # Root repository links without a branch may use master or another default.
        meta=requests.get(f"https://api.github.com/repos/{owner}/{repo}", headers=headers, timeout=30)
        if meta.ok:
            branch=str(meta.json().get("default_branch") or branch)
            api=f"https://api.github.com/repos/{owner}/{repo}/git/trees/{quote(branch, safe='')}?recursive=1"
            r=requests.get(api, headers=headers, timeout=45)
    if not r.ok:
        detail=r.text[:400].replace("\n", " ")
        raise RuntimeError(f"GitHub tree fetch failed ({r.status_code}): {detail}")
    payload=r.json(); tree=payload.get("tree", [])
    base=prefix.rstrip("/")
    md_paths=[x["path"] for x in tree if x.get("type")=="blob" and Path(x.get("path", "")).suffix.lower() in {".md", ".mdx"} and (not base or x["path"]==base or x["path"].startswith(base+"/"))]
    md_paths=sorted(md_paths)
    was_truncated=bool(payload.get("truncated"))
    if len(md_paths)>max_files: md_paths=md_paths[:max_files]
    skills=[]; total=0; failed=0
    for path in md_paths:
        raw_url=f"https://raw.githubusercontent.com/{owner}/{repo}/{quote(branch, safe='')}/{quote(path, safe='/')}"
        try:
            rr=requests.get(raw_url, headers={"Authorization": f"Bearer {token}"} if token else {}, timeout=30)
        except requests.RequestException:
            failed += 1; continue
        if not rr.ok:
            failed += 1; continue
        content=rr.text
        size=len(content.encode("utf-8"))
        if total+size>max_bytes:
            break
        total += size
        skill=parse_skill_markdown(content, f"github:{owner}/{repo}/{path}")
        if not skill.get("name") or skill["name"]=="Imported skill":
            skill["name"] = Path(path).stem if Path(path).stem.lower()!="skill" else Path(path).parent.name
        skills.append(skill)
    if was_truncated or len(skills)<len(md_paths) or failed:
        skills.append({"name":f"{repo}-import-notes", "description":"Import completeness and repository warnings", "content":f"Requested repository: https://github.com/{owner}/{repo}/tree/{branch}/{prefix}\nMarkdown paths discovered: {len(md_paths)}\nMarkdown files downloaded: {len(skills)}\nFailed downloads: {failed}\nGitHub recursive tree truncated: {was_truncated}\nImport caps: max_files={max_files}, max_bytes={max_bytes}.\nIf completeness matters, rerun with a larger authenticated import or download the repository ZIP and import it.", "source":f"github:{owner}/{repo}:import-metadata"})
    return skills


def load_bundled_skills(skill_root: str | Path | None = None) -> list[dict[str, Any]]:
    """Load bundled first-party curated Markdown playbooks from skillpacks/."""
    root=Path(skill_root) if skill_root else Path(__file__).resolve().parent/"skillpacks"
    loaded=[]
    if not root.exists():
        return loaded
    for path in sorted(root.rglob("*.md")):
        if path.is_file() and not path.is_symlink():
            try:
                content=path.read_text(encoding="utf-8")
            except (OSError, UnicodeError):
                continue
            parsed=parse_skill_markdown(content, f"bundled:{path.relative_to(root).as_posix()}")
            if not parsed.get("name") or parsed.get("name") == "Imported skill":
                parsed["name"] = path.parent.name if path.name.lower() == "skill.md" else path.stem
            loaded.append(parsed)
    return loaded


def merge_skills(existing: list[dict[str, Any]], incoming: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Deduplicate by source path so repeated imports do not flood the library."""
    result=[]; seen=set()
    for item in list(existing or []) + list(incoming or []):
        key=(item.get("source", ""), item.get("name", ""))
        if key in seen: continue
        seen.add(key); result.append(item)
    return result


def select_relevant_skills(skills: list[dict[str, Any]], query: str, limit: int=6) -> list[dict[str, Any]]:
    terms=set(re.findall(r"[a-z0-9_+#.-]{3,}", query.lower()))
    scored=[]
    for s in skills:
        text=(s.get("name","")+" "+s.get("description","")+" "+s.get("content","")).lower()
        score=sum(1 for t in terms if t in text)
        if score: scored.append((score,s))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [s for _,s in scored[:limit]]
