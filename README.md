# Base Pair — Commercial Proposal Generator

Internal Streamlit tool. Turns an Otter.ai discovery-call transcript into a
populated, non-binding commercial proposal deck in Base Pair's standard format.

## What it does

```
transcript ─(LLM extract)→ parameters ─(you review/edit)→ ─(LLM draft)→ slide prose
          ─(+ your manual milestones)→ proposal_<customer>_<date>.pptx
```

The deck is built by **template surgery**, not from scratch:

* Boilerplate slides **2–4** (Aptamer Technology, VennPlex SELEX, Value Proposition)
  are **never touched** — they are company content that doesn't change per project.
* Only the project-specific slides are rewritten:
  * **Slide 1** — subtitle, phase label, date
  * **Slide 5** — Challenge & Strategy
  * **Slide 6** — Workflow bullets
  * **Slide 7** — title, timeline line, the four timeline step boxes, and the
    milestone/pricing table (rebuilt to however many milestones you enter)

Shapes are addressed by stable `shape_id`, so edits land precisely even though the
deck reuses shape names.

## Files

| File | Purpose |
|------|---------|
| `basepair_proposal_app.py` | the app (UI + a Streamlit-free pptx engine) |
| `proposal_template.pptx`   | the canonical deck; keep it beside the app |
| `requirements.txt`         | dependencies |
| `example_output.pptx`      | a deck built from the bundled Macoska defaults (demo) |
| `basepair_logo.png`        | *optional* — shown in the app header if present |

## Setup & run

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...   # or paste it in the sidebar at runtime
streamlit run basepair_proposal_app.py
```

## Workflow in the app

1. **Transcript** — upload the Otter `.txt` (or paste it) and click
   *Extract parameters*. The LLM fills the sidebar fields; pricing is never inferred.
2. **Parameters** — review/correct everything in the sidebar expanders.
3. **Draft slides** — click *Draft Challenge / Strategy / Workflow*, then edit the
   prose in the text areas.
4. **Timeline & milestones** — set the timeline, the four step boxes, and the
   milestone rows (name / price / terms). The table resizes to your row count.
5. **Build .pptx** — download the deck.

*Save current to history* keeps the last six runs for one-click reload.

## Parameters

Your seven requested inputs, plus a few I added from reading the transcript→deck mapping:

| Parameter | Drives | Notes |
|-----------|--------|-------|
| Customer (short/full), PI/contact, institution type, date | Title slide | |
| Background & problem | Challenge slide draft | |
| Target + **target type** | Strategy/workflow + KD method | type ∈ Protein / Peptide epitope / Small molecule / Whole cells / Viral particle |
| Biological matrix | Strategy/workflow | e.g. urine |
| Off-targets to avoid | Counter-selection language | |
| **Existing aptamer? (yes/no + desc)** | *Big fork* — "test existing in matrix" vs "new SELEX" | the whole Phase-1 framing depends on this |
| **KD / affinity method** | Workflow validation step | auto-suggested from target type (Protein→MST/Octet, small molecule→MST, cells→flow/microscopy), editable |
| **Assay format goal** | Strategy/workflow endpoint | POC electrochemical / lateral-flow / etc. |
| **Phases included** | Workflow (Phase 1 only vs 1+2) | |
| **Timeline + 4 step boxes** | Slide 7 banner | |
| **Milestones (manual) + total + footnote** | Slide 7 table | you set all pricing |

## Switching LLM provider

The whole LLM surface is one function, `call_llm()`. It uses the Anthropic (Claude)
client. To move to a different provider (e.g. OpenAI), replace only the
body of `call_llm()` with that provider's call that returns the assistant
text — `llm_extract` / `llm_draft` and everything downstream are unchanged.

## Updating the boilerplate

If the company tech slides change, just drop a new `proposal_template.pptx` in place
(keeping slides 1, 5, 6, 7 as the variable slots). If you re-order slides or change
the variable shapes, update the `SLIDE_*` / `ID_*` constants at the top of the app.
