"""docforge command line."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from . import __version__, research, template
from .render import render

PROMPT_SKELETON = """\
# Instructions for the document

Template: example-report.docx

<!-- Describe in your own words what document you want. For example:
     - which template to use (the line above, or mention its file name anywhere)
     - what the document is about and who will read it
     - how the files in research/ should be used (which are most important, what to ignore)
     - language, tone, length
     Lines inside this comment are ignored. -->
"""


def _templates(project: Path) -> list[Path]:
    folder = project / "templates"
    if not folder.is_dir():
        return []
    return sorted(
        p for p in folder.iterdir()
        if p.suffix.lower() in template.TEMPLATE_SUFFIXES and not p.name.startswith("~$")
    )


def _resolve_template(project: Path, name: str) -> Path:
    for candidate in (Path(name), project / name, project / "templates" / name):
        if candidate.is_file():
            return candidate
    raise SystemExit(f"Template not found: {name}")


def _read_prompt(args, project: Path) -> str:
    if args.message:
        return args.message
    path = Path(args.prompt) if args.prompt else project / "prompt.md"
    if not path.is_file():
        raise SystemExit(f"Prompt file not found: {path}\nCreate it, or pass the instructions with -m \"...\".")
    text = path.read_text(encoding="utf-8-sig")
    text = re.sub(r"<!--.*?-->", "", text, flags=re.S).strip()
    if not text:
        raise SystemExit(f"Prompt file is empty: {path}")
    return text


def _mentioned(templates: list[Path], prompt: str) -> list[Path]:
    low = prompt.lower()
    by_name = [t for t in templates if t.name.lower() in low]
    return by_name or [t for t in templates if re.search(rf"(?<!\w){re.escape(t.stem.lower())}(?!\w)", low)]


def cmd_init(args) -> int:
    project = Path(args.folder)
    for sub in ("templates", "research", "output"):
        (project / sub).mkdir(parents=True, exist_ok=True)
    prompt = project / "prompt.md"
    if not prompt.exists():
        prompt.write_text(PROMPT_SKELETON, encoding="utf-8")
    if args.sample:
        from .sample import write_sample

        write_sample(project)
    print(f"Project ready in {project.resolve()}")
    print("  templates/  put your Word templates (.docx/.dotx) here")
    print("  research/   put your research files here")
    print("  prompt.md   describe the document you want")
    print("Then run:  docforge build")
    return 0


def cmd_inspect(args) -> int:
    project = Path(args.project)
    paths = [_resolve_template(project, args.template)] if args.template else _templates(project)
    if not paths:
        raise SystemExit(f"No templates found in {project / 'templates'}")
    for path in paths:
        doc = template.open_template(path)
        sections = template.parse_sections(doc)
        print(f"\n{path.name}: {len(sections)} sections")
        for s in sections:
            print(f"{'  ' * s.level}[{s.id}] {s.title}")
            for line in s.guidance(doc).splitlines():
                print(f"{'  ' * s.level}      | {line[:110]}")
        placeholders = template.find_placeholders(doc)
        if placeholders:
            print("  Placeholders: " + ", ".join(placeholders))
        if not sections:
            print("  No headings found. Sections are detected from Word heading styles (Heading 1-9).")
    return 0


def cmd_build(args) -> int:
    project = Path(args.project)
    out_dir = project / "output"

    if args.from_json:
        saved = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
        template_path = _resolve_template(project, args.template or saved.get("_template", ""))
        return _finish(args, template_path, saved, out_dir, save_json=False)

    prompt = _read_prompt(args, project)
    templates = _templates(project)
    model = args.model
    client = None

    if args.template:
        template_path = _resolve_template(project, args.template)
    elif not templates:
        raise SystemExit(f"No templates found in {project / 'templates'}")
    elif len(templates) == 1:
        template_path = templates[0]
    elif len(_mentioned(templates, prompt)) == 1:
        template_path = _mentioned(templates, prompt)[0]
    elif args.dry_run:
        print("Template: not named in the prompt; the model would choose between: "
              + ", ".join(t.name for t in templates))
        template_path = None
    else:
        from . import llm

        client = llm.make_client()
        print("Choosing template...")
        descriptions = {t.name: template.describe(template.open_template(t), t.name) for t in templates}
        try:
            chosen = llm.choose_template(client, model, descriptions, prompt)
        except llm.LLMError as exc:
            raise SystemExit(str(exc))
        template_path = next(t for t in templates if t.name == chosen)

    blocks, loaded, skipped = research.load(project / "research")
    if template_path is not None:
        print(f"Template: {template_path.name}")
        description = template.describe(template.open_template(template_path), template_path.name)
    print(f"Research: {len(loaded)} file(s)" + (": " + ", ".join(loaded) if loaded else ""))
    for item in skipped:
        print(f"  skipped {item}")

    if args.dry_run:
        if template_path is not None:
            print("\n" + description)
        print(f"\nInstructions:\n{prompt}\n\nDry run: nothing was sent to the model.")
        return 0

    from . import llm

    client = client or llm.make_client()
    print(f"Writing document with {model} (effort {args.effort}); this can take several minutes.")
    try:
        result = llm.write_document(
            client, model=model, effort=args.effort, max_tokens=args.max_tokens,
            research_blocks=blocks, template_description=description, prompt=prompt,
        )
    except llm.LLMError as exc:
        raise SystemExit(str(exc))
    result["_template"] = template_path.name
    return _finish(args, template_path, result, out_dir, save_json=True)


def _finish(args, template_path: Path, result: dict, out_dir: Path, save_json: bool) -> int:
    if args.output:
        out_path = Path(args.output)
        if out_path.suffix.lower() != ".docx":
            out_path = out_path.with_suffix(".docx")
    else:
        out_path = out_dir / f"{template_path.stem}-{datetime.now():%Y%m%d-%H%M%S}.docx"
    if save_json:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.with_suffix(".json").write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    try:
        stats = render(template_path, result, out_path, args.body_style)
    except PermissionError:
        raise SystemExit(f"Cannot write {out_path} - is it open in Word?")

    print(f"\nSections written: {stats['replaced']}, kept from template: {stats['kept']}")
    if stats["untouched"]:
        print("Sections the model did not return (left as in the template): " + "; ".join(stats["untouched"]))
    if stats["unknown"]:
        print("Ignored unknown section ids from the model: " + ", ".join(map(str, stats["unknown"])))
    notes = result.get("notes") or []
    if notes:
        print("\nNotes from the model:")
        for note in notes:
            print(f"  - {note}")
    usage = result.get("_usage")
    if usage and save_json:
        print(f"\nTokens: {usage['input_tokens']:,} in ({usage['cache_read_input_tokens']:,} from cache), "
              f"{usage['output_tokens']:,} out, model {usage['model']}")
    print(f"\nDocument: {out_path.resolve()}")
    return 0


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        prog="docforge",
        description="Generate Word documents from Word templates, research files and an LLM prompt.",
    )
    parser.add_argument("--version", action="version", version=f"docforge {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("init", help="create a project folder")
    p.add_argument("folder", nargs="?", default=".")
    p.add_argument("--sample", action="store_true", help="also add an example template and research file")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("inspect", help="show the sections, guidance text and placeholders found in templates")
    p.add_argument("template", nargs="?", help="template file (default: all in templates/)")
    p.add_argument("-C", "--project", default=".", help="project folder (default: current folder)")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("build", help="write the document")
    p.add_argument("-C", "--project", default=".", help="project folder (default: current folder)")
    p.add_argument("-p", "--prompt", help="prompt file (default: prompt.md in the project folder)")
    p.add_argument("-m", "--message", help="give the instructions directly instead of a prompt file")
    p.add_argument("-t", "--template", help="template to use (default: taken from the prompt)")
    p.add_argument("-o", "--output", help="output .docx (default: output/<template>-<timestamp>.docx)")
    p.add_argument("--model", default=os.environ.get("DOCFORGE_MODEL", "claude-opus-5-5"))
    p.add_argument("--effort", default="high", choices=["low", "medium", "high", "xhigh", "max"])
    p.add_argument("--max-tokens", type=int, default=64000, help="output limit for the model (default 64000)")
    p.add_argument("--body-style", help="Word paragraph style for generated body text (default: detected)")
    p.add_argument("--dry-run", action="store_true", help="show what would be sent, without calling the model")
    p.add_argument("--from-json", help="re-render a saved output/*.json without calling the model")
    p.set_defaults(func=cmd_build)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
