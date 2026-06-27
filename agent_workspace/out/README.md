# Agent deck outputs

All `.pptx` decks produced by the harness (`run_agent.py`), grouped **by model → by spec**.
Each deck has a sibling `*.log.txt` with the agent's build narrative plus a "SANDBOX CONTEXT"
note recording what was in its now-deleted sandbox.

## Specs — what each deck reproduces

Each spec is a content brief modeled on a real, publicly available source deck. The agent rebuilds
that deck from the spec; the links below are the originals it was modeled on.

- **ford-fseries** — BCG, _"The Economic Impact of Ford and the F-Series"_ (September 2020), ~21 slides.
  Source: [BCG PDF](https://web-assets.bcg.com/7d/5a/e954da90470ea56dce1472a74abd/bcg-the-economic-impact-of-ford-and-the-f-series-sept-2020-v2.pdf)
  · [BCG publication page](https://www.bcg.com/publications/2020/economic-impact-of-worlds-best-selling-truck)
- **ghp-stoptb** — McKinsey, _"Helping Global Health Partnerships to increase their impact"_ (Stop TB
  Partnership pre-reading, November 5 2009), 24 slides.
  Source: [Stop TB Partnership PDF](https://www.stoptb.org/sites/default/files/imported/document/BoardDocs/17/2.09-03_Strengthening_Performance_Management_in_the_Partnership/2.09-03.1_Helping_Global_Health_Partnerships_to_Increase_their_Impact_0.pdf)
  · [Stop TB landing page](https://www.stoptb.org/209-031-helping-global-health-partnerships-increase-their-impact)
- **ttwo** — Take-Two Interactive financial overview, ~14–15 slides (synthetic spec, no single public source deck).

## glm-5.2

> **Every deck in this section is a `z-ai/glm-5.2` run** — including all `ford-fseries` runs and
> the `ghp_run-ford3` deck (its sandbox was misnamed `out-ford3`, but the model was still glm-5.2 and
> the deck it built is the GHP/Stop-TB spec, not Ford).

| Spec                           | Deck                                              | Notes                                                                 |
| ------------------------------ | ------------------------------------------------- | --------------------------------------------------------------------- |
| ford-fseries _(BCG / Ford)_    | `ford_run1.pptx`, `ford_run2.pptx`                | later batch (~14:30)                                                  |
| ford-fseries _(BCG / Ford)_    | `collated-earlier-batch/ford_run{1,2,3}.pptx`     | earlier 3-run batch (~13:24), all ok=True / 0 problems                |
| ghp-stoptb _(McKinsey)_        | `ghp_run1.pptx`, `ghp_run2.pptx`, `ghp_run3.pptx` | three default runs                                                    |
| ghp-stoptb _(McKinsey)_        | `ghp_run-ford3.pptx`                              | sandbox was misnamed `out-ford3`, but this is a glm-5.2 GHP run       |
| ghp-stoptb _(McKinsey)_        | `ghp_traced.pptx`                                 | traced run + `*.usage.json` + `*.trace.jsonl` (was `out-ford_traced`) |
| ghp-stoptb _(McKinsey)_        | `ghp_skill.pptx`, `ghp_noskill.pptx`              | with- vs without- pptx-skill variants                                 |
| ttwo _(Take-Two)_              | `ttwo_overview.pptx`                              |                                                                       |

## gpt-5.5 (default model)

> Every deck in this section is the `run_agent.py` default model (`z-ai/gpt-5.5`).

| Spec                        | Deck                                          | Notes                                                                                 |
| --------------------------- | --------------------------------------------- | ------------------------------------------------------------------------------------- |
| ford-fseries _(BCG / Ford)_ | `ford_100pct.pptx`                            | scored 100% on /pptx-grader; file was mis-saved as `ttwo_overview.pptx` in `out-run/` |
| ttwo _(Take-Two)_           | `ttwo_overview.pptx`, `ttwo_overview_v2.pptx` | two variants                                                                          |

## Cleanup record (2026-06-26)

- Deleted all `sandbox-*/` working dirs (repo copies + uv/pip caches) and the old `out-*/` dirs.
- Some GHP decks (`ghp_run1/2/3`, `ghp_noskill`) existed only inside their sandbox and were rescued here before deletion.
- Model attribution: everything under `glm-5.2/` is `z-ai/glm-5.2`; everything under `gpt-5.5/` is the `run_agent.py` default model.
- `ghp_run2` had no run_log (its `out-ghp2/` was empty) — see its log for the reconstructed note.
