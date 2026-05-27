# Restaurant Proposal Generator Migration Pack

This package contains the minimum code and template assets needed to run the
restaurant menu analysis and cooperation PPT generation workflow on another PC.

## Included

- `proposalGen/`: restaurant menu analysis, lotus-root opportunity analysis,
  image generation wrapper, and PPT proposal generator scripts.
- `examples/sample.pptx`: the source PowerPoint template used by the proposal
  generator.
- `skills/ppt-master/scripts/image_gen.py`: shared image-generation runner.
- `skills/ppt-master/scripts/image_backends/`: image backend helpers.
- `skills/ppt-master/scripts/config.py`, `error_helper.py`,
  `gemini_watermark_remover.py`: helper modules required by `image_gen.py`.
- `requirements.txt` and `skills/ppt-master/requirements.txt`: Python package
  dependencies.

## Not Included

- Real `.env` files or API keys.
- Generated project outputs under `projects/`.

Copy your own `.env` values from the old machine, or fill in the provided
`.env.example` and rename it to `.env`.

## Setup On The New PC

Run these commands from the package root:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Then edit `.env` and fill in the API credentials and model names.

## Typical Workflow

```powershell
python -B proposalGen/restaurant_menu_analyzer.py "Restaurant Name"
python -B proposalGen/lotus_root_opportunity_analyzer.py projects\menu_research\<run_dir>\structured_menu.json
python -B proposalGen/lotus_root_image_generator.py projects\menu_research\<run_dir>\lotus_root_opportunities.json
python -B proposalGen/create_restaurant_coop_ppt.py --project-dir projects\menu_research\<run_dir> --restaurant-name "Restaurant Name"
```

For an existing project directory copied from the old machine, start from the
step that matches the files already present in that project directory.
