# AI Tools Quick Start

_Two worked examples: an excise diagnostic for Indonesia (Claude) and new tax indicators in the dashboard (GitHub Copilot). October 2026._

**Download everything:** [`AI_Tools_Quick_Start_Kit.zip`](AI_Tools_Quick_Start_Kit.zip) - this guide (Word, PowerPoint), the copy-paste prompts and every file the AI needs.

## Which AI for what

| Tool | Where | What it does |
|---|---|---|
| **Claude** | claude.ai | Deep research and writing. Finds laws and budget papers on the web, reads them, and writes the report and workbook. Use it for Example 1. |
| **GitHub Copilot** | VS Code or the Copilot app | Changes files in our repo: code, config, rebuilds the dashboard, pushes to GitHub. Use it for Example 2. |
| **.md process files** | in the repo | Plain-text instructions the AI follows step by step. You give them to the AI. You do not need to edit them. |

## Before you start (once)

- A Claude account that has Projects and Research (deep research).
- GitHub Copilot set up, and the repo cloned to C:\Users\<you>\Repos\new-ai-tools (see Presentations/GitHub_Copilot_AI_User_Guide.pptx).
- This kit unzipped. Folder 1 is for Example 1, folder 2 for Example 2.

## Example 1 - Excise diagnostic for Indonesia (Claude)

A structured snapshot of Indonesia's excise schedule at 30 Sept 2026: every excise line, quoted from the law, standardized, with gaps and erosion flagged. Output: a Word report and an Excel workbook.

1. **Create a Project.** On claude.ai: Projects > Create project. Name it "Excise - Indonesia".
2. **Add 3 files.** Project knowledge > Add: excise_diagnostic_process.md, Excise_Diagnostic_CrossCountry.xlsx, INDONESIA_INPUTS.md (all in kit folder 1).
3. **Start a chat.** New chat inside the Project. Turn on Research and Web search. Pick the most capable model.
4. **Paste Prompt 0, then A1.** From PROMPTS.md. Prompt 0 checks Claude understood. A1 builds the legal history.
5. **If Claude asks for files.** It gives you a links list and a .bat file. Run the .bat (or click the links), upload the PDFs to the Project, type "continue".
6. **Check, then next prompt.** Every rate quoted with its provision and date, every source linked. Then send A2, A3 ... A6, one per message.
7. **Save the outputs.** Download the Word report and Excel workbook into ExciseDiagnostic/country_results/Indonesia/.

The six prompts (full text in [`kit/1_Excise_Indonesia_Claude/PROMPTS.md`](kit/1_Excise_Indonesia_Claude/PROMPTS.md)):

| Prompt | What you get |
|---|---|
| A1 | Legislative history, 2020 to cut-off, rates as printed |
| A2 | Current schedule, every line, rate trajectory |
| A3 | Parameters, standard schedule, workbook |
| A4 | Findings (Level 1), every number traced |
| A5 | Peer-review register, corrections applied |
| A6 | Final Word report and Excel workbook |

> Short on time? Add excise_part1_fuel.md to the Project and run Part 1 (fuel) only, with the prompts in that file.
>
> Want a fossil fuel SUBSIDY diagnostic instead? Use the variant prompt at the end of PROMPTS.md. There is no subsidy process file yet, so treat the output as a first draft.

## Example 2 - Add tax indicators to the dashboard (GitHub Copilot)

Add four rows under Tax revenue (% GDP) in the Global Fiscal Indicators dashboard: personal income tax, corporate income tax, VAT and excises, all % of GDP. The data are already in the repo (IMF WoRLD 2026), so nothing to download.

1. **Open the repo.** VS Code > File > Open Folder > Repos\new-ai-tools. Pull (Sync) first.
2. **Open Copilot Chat.** Choose Agent mode and the model Claude Sonnet.
3. **Prompt 1 - plan.** Copilot reads the README and lists what it will change. Nothing is edited yet.
4. **Prompt 2 - edit.** It edits 3 files: the Python build, the R build and config/indicators.csv. Click Keep.
5. **Prompt 3 - build and check.** It rebuilds and shows Indonesia's rows. Compare with the numbers below.
6. **Prompt 4 - new version.** It ships v1.4 of the workbook and updates VersionNotes.md.
7. **Prompt 5 - push.** It commits and pushes to GitHub (or double-click SyncRepo.bat).

Full prompts: [`kit/2_Tax_Indicators_Copilot/PROMPTS.md`](kit/2_Tax_Indicators_Copilot/PROMPTS.md). Check: Indonesia, 2024, % of GDP:

| Row | Expected |
|---|---|
| Tax revenue | 10.08 |
| Personal income tax | 1.79 |
| Corporate income tax | 2.65 |
| VAT / general sales tax | 3.52 |
| Excises | 1.11 |

## Five golden rules

1. One prompt per message. Wait for the answer.
2. Check each answer before the next prompt. You are the reviewer.
3. No number without a source. Missing data stays missing (data_gap).
4. Ask for a plan first ("do not edit yet"), then say "go ahead".
5. Save outputs in the repo folder, then push, so the team has them.

## What is in the zip

| File | What it is |
|---|---|
| `AI_Tools_Quick_Start.docx / .pptx / README.md` | This guide |
| `1_Excise_Indonesia_Claude/PROMPTS.md` | Copy-paste prompts A1-A6, plus the subsidy variant |
| `1_Excise_Indonesia_Claude/INDONESIA_INPUTS.md` | Country, currency, cut-off, the laws to look for |
| `1_Excise_Indonesia_Claude/process_files/` | The six .md process files (master + 5 parts) |
| `1_Excise_Indonesia_Claude/Excise_Diagnostic_CrossCountry.xlsx` | Workbook layout for Claude to follow |
| `1_Excise_Indonesia_Claude/Example_SLE_Excise_Diagnostic.docx` | What a finished report looks like (Sierra Leone) |
| `2_Tax_Indicators_Copilot/PROMPTS.md` | Copy-paste prompts 1-5 and the checks |

Rebuild this guide, the Word and PowerPoint versions and the zip with `python user_guide/build_user_guide.py`.
