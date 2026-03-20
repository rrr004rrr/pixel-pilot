# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

GUI automation tool that uses image template matching (OpenCV) to find and click UI elements whose positions change dynamically — fixed coordinates cannot be used.

## Commands

```bash
pip install -r requirements.txt           # install dependencies (once)
python workflow.py                         # run workflow from step 1
python workflow.py 5                       # resume from step 5
python auto_clicker.py templates/btn.png --no-click           # test if template is found
python auto_clicker.py templates/btn.png --no-click --debug   # test with visual overlay
python auto_clicker.py templates/btn.png --no-click --confidence 0.75  # lower threshold
```

## Architecture

Two modules with strict separation of concerns:

**`auto_clicker.py`** — low-level library. Do not modify core functions unless fixing a bug.

| Function | Behavior |
|----------|----------|
| `find_and_click(path)` | Wait until template appears, then click |
| `find_only(path)` | Return `(x, y)` without clicking |
| `image_exists(path)` | Instant boolean check |
| `wait_for_image(path)` | Block until template appears |
| `wait_for_image_gone(path)` | Block until template disappears |
| `find_all(path, click_all)` | Find all matches, optionally click each |
| `sleep(seconds, reason)` | Annotated wait |

Key parameters: `confidence` (default 0.8, fallback 0.75), `wait_timeout` (default 10s).

**`workflow.py`** — orchestration layer. Define steps here; never add try/except inside step functions.

- Each step is a standalone function named `step_NN_name()`
- All steps are registered in the `STEPS` list with a failure behavior per step:
  - `"stop"` — halt the workflow
  - `"skip"` — log and continue
  - `"retry"` — retry once, then continue
- The `run()` engine iterates `STEPS`; do not modify it

## Adding a New Step

1. Put the button screenshot (PNG) in `templates/`
2. Add `def step_NN_name()` in `workflow.py` using `auto_clicker` functions
3. Append `{"func": step_NN_name, "on_fail": "stop"}` to `STEPS`

## Template Screenshot Rules

- Crop tightly to the button/icon — no extra background
- PNG format, captured at the same screen resolution as runtime (no scaling)
- Filename: lowercase English with underscores, e.g. `btn_confirm.png`, `icon_loading.png`

## Safety

- Moving the mouse to the **top-left corner** immediately stops execution (PyAutoGUI FailSafe)
- 0.3-second pause is applied after every operation automatically
