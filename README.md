# TaraGO (タラゴー)

**Tara! + 語** — an evidence-based, multi-skill Japanese learning app for Filipino fresh graduates targeting employment in Japan (SSW / Engineer / general N5→N2 tracks).

## Project structure

| Path | Contents |
|---|---|
| `docs/TaraGO_Thesis_Proposal.docx` | Mapúa IEEE-style thesis proposal (15 pp, 32 references) justifying the app. PDF copy included. |
| `docs/TaraGO_Study_Pack.docx` | Printable study pack: 12-week N5→N4 plan, 109 kanji, 349 vocab, 40 grammar points, writing & speaking guides. PDF copy included. |
| `app/TaraGO_N5_Reviewer.html` | Interactive N5 reviewer prototype — open in Chrome/Edge. SRS flashcards, quizzes, kanji/vocab/grammar/kana reference, shadowing with TTS + mic check, handwriting canvas. Progress saves in browser localStorage. |
| `scripts/` | Node scripts that regenerate the two Word documents (`npm i docx`, then `node make_thesis.js` / `node make_pack.js`). `make_pack.js` reads its data from the reviewer HTML — the HTML is the single source of truth for content. |

## Evidence base (core claims)

- Spaced repetition + retrieval practice → retention (Cepeda et al. 2006; Roediger & Karpicke 2006)
- Kanji via component analysis + light mnemonics (Rose 2013; Gamage 2003)
- Speaking via shadowing + ASR feedback, g = 0.69 (Hamada 2019; Ngo, Chen & Lai 2024)
- Handwriting production for character memory (Longcamp et al. 2005; Naka & Naoi 1995)
- PH context: 44,457 learners (Japan Foundation 2021); instrumental motivation (Gonzales); JPEPA ~7% early pass rates show language is the gating skill

## Roadmap

1. ✅ Thesis proposal + N5 reviewer prototype + study pack
2. ⬜ N4 content wave (kanji 101–300, N4 grammar, JFT-Basic mocks)
3. ⬜ Mobile app build (Flutter) with real SRS scheduler + offline packs
4. ⬜ ASR pronunciation scoring; occupational modules (SSW caregiving/food service, IT/keigo)
5. ⬜ Pilot evaluation per Chapter 3 (ISO 25010 + quasi-experiment)

## Git

Repo owner: **newtonsuu** (kingjerichojames@gmail.com). Commit style: frequent, per-milestone commits.
