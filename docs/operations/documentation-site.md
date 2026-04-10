# Documentation Site

## What was added

This repository now includes a Read the Docs compatible documentation setup using MkDocs.

Important files:

- `.readthedocs.yaml`
- `mkdocs.yml`
- `docs/`
- `docs/requirements.txt`

## Local docs workflow

Install docs dependencies:

```bash
python3 -m pip install -r docs/requirements.txt
```

Serve locally:

```bash
mkdocs serve
```

Build locally:

```bash
mkdocs build --strict
```

## Read the Docs publishing

The included `.readthedocs.yaml` tells Read the Docs to:

- use Python 3.12
- install `docs/requirements.txt`
- build the site using `mkdocs.yml`

## Documentation maintenance rules

When architecture changes, update docs alongside code changes in the same branch whenever possible.

The most important pages to keep current are:

- repository layout
- browser runtime
- desktop bridge
- configuration
- build and deployment

## Recommended habit

If you move folders, rename shared runtime files, or change sync flows, update the docs before the next release so the repository structure and docs never drift apart.
