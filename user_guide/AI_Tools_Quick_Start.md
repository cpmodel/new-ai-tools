# AI Tools Quick Start (GitHub Copilot)

**You need:** the new-ai-tools repo on your laptop, and GitHub Copilot in VS Code.

**To start:** open the repo in VS Code, click **Sync**, open **Copilot Chat** in **Agent** mode.
Paste one prompt. Read the answer. Then paste the next.

---

## 1. Excise diagnostic for Indonesia

Model: **Claude Opus**. **⚠ Uses a lot of credits: one prompt at a time.**

The instructions Copilot follows: [`excise_diagnostic_process.md`](excise_diagnostic_process.md) (in this folder).

```text
Read user_guide/excise_diagnostic_process.md. Run stage A1 for Indonesia, 2020 to 2026-09-30.
Stop after A1.
```

Then, one at a time: `Run stage A2.` · `Run stage A3.` · `Run stage A4.` · `Run stage A5.` ·
`Run stage A6 and save the Word report and Excel workbook in ExciseDiagnostic/country_results/Indonesia/.`

**Check:** every rate has a date and a source link.

---

## 2. Add tax indicators to the dashboard

Model: **Claude Sonnet**.

```text
Add VAT, personal income tax, corporate income tax and excises (% of GDP) to the dashboard,
under Tax revenue, from the IMF WoRLD file already in raw_data.
Follow TaxMacroFiscalDashboard/build/README.md. Then rebuild and show me Indonesia.
```

**Check (Indonesia, % of GDP):** tax 10.08 · VAT 3.52 · corporate 2.65 · personal 1.79 · excises 1.11

Then: `Commit and push.`

---

**Appendix – with Claude instead:** make a Claude Project, add `excise_diagnostic_process.md`, switch on Research, and paste the same prompts.
