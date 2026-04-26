# Equity Research — `@modestproposal1` Bloomberg Screenshot Archiver

A small Python tool that walks the recent timeline of
[`@modestproposal1`](https://x.com/modestproposal1) on X (Twitter), downloads
the photos he attaches to his tweets, and keeps the ones that look like
Bloomberg Terminal or Bloomberg News screenshots.

`@modestproposal1` is an anonymous investor whose feed is dotted with
Bloomberg charts and headlines — useful raw material for equity-research
note-taking. Rather than scrolling X by hand, this tool pulls the images
straight from the API and persists a JSONL index alongside the cropped
screenshots.

> **Heads up:** the X API requires authenticated access. You need a
> developer account and a Bearer Token (free tier is fine for this use
> case). Without credentials the tool will refuse to run.

## How it works

1. **Fetch** — `src/modestproposal_bloomberg/fetcher.py` calls the X v2
   `users/:id/tweets` endpoint with `media.fields` expanded and yields a
   stream of `TweetMedia` objects (one per attached photo).
2. **Classify** — `src/modestproposal_bloomberg/bloomberg_filter.py`
   inspects each downloaded image. Two heuristics run in parallel:
   - a pixel histogram check that looks for the dark background and amber
     text that defines a Bloomberg Terminal frame, and
   - an OCR pass (via `pytesseract`, when installed) plus tweet
     caption/alt-text scan for the literal string `Bloomberg`, `BBG`, or a
     command of the form `<GO>`.
3. **Archive** — `src/modestproposal_bloomberg/archiver.py` writes matching
   images to `data/screenshots/` and appends a metadata record to
   `data/index.jsonl`. Already-seen `media_key`s are skipped on subsequent
   runs so the tool is safe to schedule.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

OCR is optional but recommended. On Debian/Ubuntu:

```bash
sudo apt-get install -y tesseract-ocr
```

## Configuration

Copy `.env.example` to `.env` and fill in your credentials:

```bash
cp .env.example .env
$EDITOR .env
```

| Variable          | Required | Notes                                              |
| ----------------- | -------- | -------------------------------------------------- |
| `X_BEARER_TOKEN`  | yes      | App-only Bearer Token from developer.x.com.        |
| `TARGET_HANDLE`   | no       | Defaults to `modestproposal1`.                     |
| `OUTPUT_DIR`      | no       | Defaults to `./data`.                              |

## Usage

```bash
# One-off run: archive the last few pages of @modestproposal1's tweets.
modestproposal-bloomberg --max-pages 3 -v

# Restrict to tweets newer than a specific date.
modestproposal-bloomberg --since 2026-04-01

# Different handle, no OCR, save every photo (not just Bloomberg ones).
modestproposal-bloomberg --handle someone_else --no-ocr --save-non-matches
```

After a run you'll find:

```
data/
  index.jsonl          # one JSON record per evaluated photo
  last_run.json        # summary of the most recent invocation
  screenshots/         # archived Bloomberg images
```

Each `index.jsonl` row contains the source tweet metadata, the classifier
features (dark/amber/cyan ratios, OCR excerpt, matched keywords, score),
and the on-disk path of the saved screenshot when applicable.

## Running the tests

```bash
pip install pytest
PYTHONPATH=src pytest
```

The test suite covers the Bloomberg classifier on synthetic images and
exercises the X client against an in-memory fake session, so no network
or credentials are required.

## Layout

```
src/modestproposal_bloomberg/
  __init__.py            # public re-exports
  fetcher.py             # XClient + TweetMedia
  bloomberg_filter.py    # BloombergScreenshotClassifier
  archiver.py            # Archiver + ArchiveSummary
  cli.py                 # `modestproposal-bloomberg` entry point
tests/                   # pytest suite (no network)
```

## Caveats

- The classifier is a heuristic. Terminal screenshots are very reliable;
  Bloomberg.com screenshots depend on OCR catching the watermark or on
  the tweet text mentioning Bloomberg. Tune the thresholds in
  `BloombergScreenshotClassifier(...)` if you want to be more or less
  aggressive.
- The X free API tier has tight rate limits. The client honors
  `x-rate-limit-reset` and backs off on 5xx, but a long backfill will
  still take a while.
- Respect X's developer terms of service and `@modestproposal1`'s
  copyright: this tool is meant for personal research notebooks, not
  republication.
