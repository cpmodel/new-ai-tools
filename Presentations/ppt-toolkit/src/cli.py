"""ppt-toolkit CLI: render, harvest, convert.

Usage:
    ppt render deck.json --template wb -o out.pptx
    ppt harvest old.pptx -o extracted.json
    ppt convert old.pptx --template white -o new.pptx
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from render import main as render_main  # noqa: E402
from harvest import harvest, main as harvest_main  # noqa: E402
from render import render_deck  # noqa: E402


def cmd_render(args):
    argv = [args.deck_json, "--template", args.template, "-o", args.output]
    if args.skip_validate:
        argv.append("--skip-validate")
    render_main(argv)


def cmd_harvest(args):
    harvest_main([args.pptx_path, "-o", args.output])


def cmd_convert(args):
    output_dir = os.path.dirname(os.path.abspath(args.output)) or "."
    images_dir = os.path.join(output_dir, "images")
    deck = harvest(args.pptx_path, images_dir=images_dir)
    prs = render_deck(deck, args.template, output_dir)
    os.makedirs(output_dir or ".", exist_ok=True)
    prs.save(args.output)
    print(f"converted {args.pptx_path} -> {args.output} (template={args.template})")


def main():
    parser = argparse.ArgumentParser(prog="ppt", description="ppt-toolkit CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    p_render = sub.add_parser("render", help="Render a deck JSON file into a .pptx")
    p_render.add_argument("deck_json")
    p_render.add_argument("--template", choices=["white", "wb"], required=True)
    p_render.add_argument("-o", "--output", required=True)
    p_render.add_argument("--skip-validate", action="store_true")
    p_render.set_defaults(func=cmd_render)

    p_harvest = sub.add_parser("harvest", help="Extract an existing .pptx into deck JSON")
    p_harvest.add_argument("pptx_path")
    p_harvest.add_argument("-o", "--output", required=True)
    p_harvest.set_defaults(func=cmd_harvest)

    p_convert = sub.add_parser("convert", help="Harvest + render in one step (old.pptx -> new template)")
    p_convert.add_argument("pptx_path")
    p_convert.add_argument("--template", choices=["white", "wb"], required=True)
    p_convert.add_argument("-o", "--output", required=True)
    p_convert.set_defaults(func=cmd_convert)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
