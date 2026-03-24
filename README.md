# TaraGO (タラゴー)

**Tara! + 語** — an evidence-based, multi-skill Japanese learning app for Filipino fresh graduates targeting employment in Japan (SSW / Engineer / general N5→N2 tracks).

## Project structure

| Path | Contents |
|---|---|
| `docs/TaraGO_Thesis_Proposal.docx` | Mapúa IEEE-style thesis proposal (15 pp, 32 references) justifying the app. PDF copy included. |
| `docs/TaraGO_Study_Pack.docx` | Printable study pack: 12-week N5→N4 plan, 109 kanji, 349 vocab, 40 grammar points, writing & speaking guides. PDF copy included. |
| `index.html` | The TaraGO web app (PWA). **Learn**: Duolingo-style N5→N1 path — units with progress rings, mixed-exercise lessons (MCQ, listening, match-pairs, word-order), XP/streak/daily goal. **Review**: SRS flashcards. **Mock Exam**: JFT-Basic style. **AI Sensei**: Claude-powered tutor (bring your own API key). **Library**: kanji/vocab/grammar/kana reference, shadowing with TTS + mic check, handwriting canvas. Progress saves in browser localStorage. |
| `scripts/` | Node scripts that regenerate the two Word documents (`npm i docx`, then `node make_thesis.js` / `node make_pack.js`). `make_pack.js` reads its data from the reviewer HTML — the HTML is the single source of truth for content. |

## Evidence base (core claims)

- Spaced repetition + retrieval practice → retention (Cepeda et al. 2006; Roediger & Karpicke 2006)
- Kanji via component analysis + light mnemonics (Rose 2013; Gamage 2003)
- Speaking via shadowing + ASR feedback, g = 0.69 (Hamada 2019; Ngo, Chen & Lai 2024)
- Handwriting production for character memory (Longcamp et al. 2005; Naka & Naoi 1995)
- PH context: 44,457 learners (Japan Foundation 2021); instrumental motivation (Gonzales); JPEPA ~7% early pass rates show language is the gating skill

## Roadmap

1. ✅ Thesis proposal + study pack + PWA
2. ✅ N4 content wave (163 kanji, 40 grammar, JFT-Basic mock exam)
3. ✅ Duolingo-style learn path N5→N1: units, progress rings, lesson engine, XP/streak/daily goal
4. ✅ N3/N2/N1 starter content waves + AI Sensei (Claude tutor: grammar explanations, sentence checking, interview roleplay)
5. ⬜ Content completion waves toward full JLPT coverage (N3 ≈ 650 kanji cum., N2 ≈ 1,000, N1 ≈ 2,136; vocab to ~10k)
6. ⬜ Occupational modules (SSW caregiving/food service, IT/keigo) as path overlays
7. ⬜ ASR pronunciation scoring beyond browser speech recognition
8. ⬜ Pilot evaluation per Chapter 3 (ISO 25010 + quasi-experiment)

## Deployment

Hosted on GitHub Pages from `main` branch root. On iPhone: open the site in Safari → Share → **Add to Home Screen** → launches full-screen and works offline (service worker).

## Git

Repo owner: **newtonsuu** (kingjerichojames@gmail.com). Commit style: frequent, per-milestone commits.
