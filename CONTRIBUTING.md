# Contributing to TabEscape

Small, focused changes are welcome. Issues and pull requests can be in English or Russian.

## Work on the running application

The canonical app lives in `tabescape_app/`. `TabEscape (2).html` is a historical standalone design reference; editing it does not update the Flask website.

1. Follow the [local setup](README.md#установка-и-запуск).
2. Create a branch such as `codex/fix-workbook-import`.
3. Keep the shared RPG state and `/api/start` / `/api/action` contract intact. A mask is a presentation layer, not a new game engine.
4. After changing templates or static files, run `python scripts/embed_assets.py` from `tabescape_app`. Commit its `app.py` update as well.
5. Run `python -m unittest discover -s tests -v` and `gunicorn app:app --check-config --timeout 60` from that directory.
6. For UI changes, check the affected mask at desktop and narrow widths. Make a turn, switch with Esc and the on-screen button, then return to the same scene. Record what you actually checked.

CI runs Python 3.11/3.12 tests, JavaScript syntax checks, Gunicorn configuration, embedded-asset consistency and dependency auditing. Tests use mocked OpenAI responses; contributors do not need a paid API key.

## Adding a campaign

Add a RU/EN entry in `static/data/campaigns.json` with a unique ID, genre, start node, choices, deltas and endings. Every choice must point to an existing node or ending. Keep choices understandable, consequences tied to the fiction and scenes short enough for a work break. Update catalogue tests, README and the embedded resources.

## Credentials and data

Keep local keys in `tabescape_app/.env`, which Git ignores. Never paste keys or private workbooks into issues, screenshots or logs. Both committed `.env.example` files contain placeholders only. Avoid enabling live generation in tests or CI.

## Scope and attribution

This repository is a hackathon demo. Do not claim payment processing, durable saves, live-AI validation or full Office/DWG compatibility unless the change actually implements and verifies them. Describe limitations in the PR.

No open-source license has been selected yet. Do not add a license or third-party artwork on behalf of the team without establishing the appropriate rights. See [asset provenance](docs/ASSETS.md).
