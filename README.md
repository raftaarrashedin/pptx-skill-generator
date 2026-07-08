# Presentation Builder (v1)

Script-based pipeline: template → request → draft (deepagents) → review →
revise (deepagents) → render (SVG) → export (pptx). Everything reads/writes
`output/` only. One script = one stage, each readable top to bottom.

Note: your original folder sketch labeled the output dir `utils/`, but every
script reference in the plan text said `output/`. I went with `output/`.
Rename in each script's `--output-dir` default if you actually meant `utils/`.

## Setup

```bash
pip install -r requirements.txt
# Optional only if your OpenAI-compatible endpoint expects a real key.
export OPENAI_API_KEY=your-key
```

Live LLM calls now go through one shared entry point in `llm_client.py`,
using the active registry entry `openai/qwen3.6:35b` via the configured
OpenAI-compatible base URL. `--mock` still forces deterministic stub
content so you can test the pipeline plumbing without hitting the model.

## Run it

Saved templates can live in the project root or in `templates/`. You can then
refer to them by name instead of a full path, for example `--template "Nature Journal"`.

One command, default flow, mock content, stops for your review:

```bash
python3 run_presentation_builder.py all \
  --title "Q3 Board Update" --topic "Quarterly performance review" \
  --slides 3 --audience "Executive leadership"
```

Check `output/review.txt`, then continue:

```bash
python run_presentation_builder.py all --title "..." --topic "..." --auto-approve
# or, with the configured live Qwen endpoint instead of mock content:
python run_presentation_builder.py all --title "..." --topic "..." --live --auto-approve
```

### Stage by stage

```bash
python run_presentation_builder.py templates
python run_presentation_builder.py template [--template path.pptx]
python run_presentation_builder.py template --template "Nature Journal"
python run_presentation_builder.py request --title "..." --topic "..." --slides N
                                  # (no flags = interactive prompts)
python run_presentation_builder.py orchestrate [--feedback "..."] [--mock]
python run_presentation_builder.py review [--approve]
python run_presentation_builder.py revise --feedback "..." [--mock]
python run_presentation_builder.py render
python run_presentation_builder.py export
python run_presentation_builder.py edit move-slide --id 4 --before 2
python run_presentation_builder.py edit edit-title --id 3 --title "New Title"
python run_presentation_builder.py edit remove-bullet --id 2 --index 1
python run_presentation_builder.py edit duplicate-slide --id 3
python run_presentation_builder.py edit delete-slide --id 5
```

Every script also runs standalone (`python generate_presentation_draft.py --mock`, etc.) — the runner is a thin dispatcher, not the only way in.

Example with a saved template name:

```bash
python3 run_presentation_builder.py all \
  --template "Nature Journal" \
  --title "Q3 Board Update" \
  --topic "Artificial Intelligence" \
  --slides 10 \
  --audience "Executive leadership" \
  --live \
  --auto-approve
```

## Files

```
output/
  template_context.json     # design context: fonts, colors, slide size, extracted assets
  presentation_request.json # structured user intent
  presentation_draft.json   # source of truth, slide-by-slide, under revision
  presentation_final.json   # approved snapshot of the draft, used by render/export
  preview.json               # slide counts/types + content flags
  review.txt                 # human-readable outline
  slides_svg/slide_NN.svg    # deterministic layout per slide, data-role tagged
  final_presentation.pptx    # editable output (native text boxes + native chart objects)
  assets/                     # images extracted from a source template, if given
  run.log
```

## Design notes

- **`presentation_draft.json` is the source of truth**, not the `.pptx`. Every
  stage after generation reads/writes this JSON (or its `_final` snapshot).
- **Only generation, revision, and orchestration touch an LLM**. They all go
  through `llm_client.py`, which centralizes the active model config and the
  `deepagents` invocation path so output shape stays consistent. Render and
  export are deterministic.
- **The SVG is a narrow, renderer-specific dialect**, not general SVG — every
  element carries a `data-role` (and `data-category` / `data-value` / etc.)
  that `export_presentation_pptx.py` reads back to rebuild native shapes.
  Don't feed it arbitrary SVG.
- **Chart slides export as real, editable PowerPoint chart objects**
  (`add_chart`, category data), not flattened images — double-click to edit
  in PowerPoint/Excel.
- Slide types: `title`, `section`, `bullets`, `metrics`, `comparison`,
  `chart`, `image_text`.
- `edit_presentation.py` does not re-render automatically — run `render` +
  `export` again after editing to see changes in the `.pptx`.
