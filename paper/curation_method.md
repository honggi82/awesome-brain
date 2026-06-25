# Awesome Brain Curation Method

Generated: 2026-06-26

## Scope

- Topic: brain research
- Years: 1900-2026
- Candidate target: up to 1,000 papers per year
- Selection target: 100 papers per year
- Ranking: citation count descending, using Semantic Scholar `citationCount`
- Metadata source: Semantic Scholar Academic Graph bulk search, free public metadata

## Query

For each year, the pipeline sends broad brain and neuroscience queries to Semantic Scholar, sorts by citation count, and paginates within the free public endpoint limits. Local relevance checks then keep records whose title, abstract, fields, or publication metadata indicate brain/neuroscience content.

For early years where public metadata contains fewer than the requested target, the generated datasets keep the full audited pool and select all available citation-ranked records rather than fabricating missing entries.

## Enrichment

The script deterministically assigns taxonomy categories, keyword convention tags, key ideas, strengths, and research-focused limitations. No paid API, paid LLM, paid translation service, or paid compute is used.

## Verification Targets

The repository should contain selected and candidate CSV/JSON data, `README.md`, `README.html`, `docs/index.html`, period analysis JSON, taxonomy SVG assets, and English/Korean review HTML files.
