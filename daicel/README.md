# Gantt Chart Generator

A browser-based Jira CSV to Gantt chart tool.

## Features

- Upload one or more Jira CSV exports
- Render an editable Gantt chart in the browser
- Filter by project
- Group by type, parent, or no grouping
- Export PNG from the browser
- Export PPTX from the browser on static hosting
- Export PPTX/PNG through `server.py` when running locally

## Static Web Version

This folder can be published directly with GitHub Pages. The main entry point is:

```text
index.html
```

For GitHub Pages, enable Pages from the repository settings and publish from the root of the default branch.

## Local Server Version

If Python and `python-pptx` are available, run:

```bash
python server.py
```

Then open:

```text
http://localhost:8090/
```

Local exports are saved under:

```text
output/
```
