# MacroFiscalModels/Docs: pitch deck and user guides

| File | What it is |
|---|---|
| `MacroFiscalModels_Pitch_v0_1.pptx` | 11-slide pitch: the question, the toolkit, the NMIM link, Kenya results (GIMF-lite, DIGNAD-lite, LIC-DSF-lite), data needs, status, next steps |
| `UserGuide_0_GettingStarted.docx` / `.md` | Install, run the demo, run your own scenario, how the pieces connect, rules for using the results |
| `UserGuide_1_GIMF_lite.docx` / `.md` | GIMF-lite: parameters (with data sources), outputs, recipes, reading the results, limits |
| `UserGuide_2_DIGNAD_lite.docx` / `.md` | DIGNAD-lite: the same for disasters, adaptation and public investment financing |
| `UserGuide_3_LIC_DSF_lite.docx` / `.md` | LIC-DSF-lite: capacity, inputs, indicators, stress tests, rating |
| `UserGuide_4_Integrated.docx` / `.md` | Integrated country run: the chain, model choice, outputs, dashboard Appendix D, caveats |

The Markdown files are the sources; the Word files and the deck are generated:

```bash
cd MacroFiscalModels
python3 Docs/export_deck_data.py KEN      # numbers for the deck -> Docs/deck_data.json
cd Docs
node build_pitch.js                       # needs pptxgenjs, react, react-dom, react-icons, sharp
./build_docs.sh                           # needs pandoc; styles from the investment needs note
```

Edit the Markdown or `build_pitch.js`, not the Word or PowerPoint files, so they stay in step.
