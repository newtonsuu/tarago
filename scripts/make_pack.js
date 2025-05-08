const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, PageBreak, LevelFormat
} = require("docx");
const fs = require("fs");

// ---- extract data from the reviewer HTML (single source of truth) ----
const html = fs.readFileSync("TaraGO_N5_Reviewer.html", "utf8");
const js = html.match(/<script>([\s\S]*)<\/script>/)[1];
const KANJI_RAW = js.match(/const KANJI_RAW=`([\s\S]*?)`;/)[1];
const VOCAB_RAW = js.match(/const VOCAB_RAW=`([\s\S]*?)`;/)[1];
const GRAMMAR = eval(js.match(/const GRAMMAR=(\[[\s\S]*?\]\]);/)[1]);
const KANJI = KANJI_RAW.trim().split("\n").map(l => l.split("|")).filter(a => a.length === 4 && a[0].length === 1);
KANJI.push(["西", "セイ・サイ", "にし", "west"]);
const VOCAB = VOCAB_RAW.trim().split("\n").map(l => l.split("|")).filter(a => a.length === 4);

const LATIN = "Calibri", JP = "Yu Gothic";
function P(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    spacing: { after: o.after ?? 120, line: 276 },
    children: [new TextRun({ text, font: o.jp ? JP : LATIN, size: o.size || 22, bold: o.bold, italics: o.italics, color: o.color })],
  });
}
function H1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 160 }, children: [new TextRun({ text: t, font: LATIN, size: 32, bold: true, color: "C0392B" })] }); }
function H2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 }, children: [new TextRun({ text: t, font: LATIN, size: 26, bold: true, color: "2C3E50" })] }); }
function cell(t, w, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: opts.head ? { type: ShadingType.CLEAR, fill: "F2D7D5" } : undefined,
    children: [new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: t, font: opts.jpFont ? JP : LATIN, size: opts.size || 20, bold: opts.head })] })],
  });
}
function tbl(widths, headTexts, rows) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ children: headTexts.map((h, i) => cell(h, widths[i], { head: true })) }),
      ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, widths[i], { jpFont: i === 0 || /[぀-ヿ一-鿿]/.test(c), size: 20 })) })),
    ],
  });
}
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ---- 12-week plan data ----
const PLAN = [
["1", "N5 reset: kana + core kanji", "Re-read kana charts until instant recall; SRS: first 30 kanji; 50 vocab; grammar は/です/も/の", "Flashcards 20 min/day; write each kanji 5× by hand; shadow drills 1–5"],
["2", "N5 kanji block 2", "SRS: kanji 31–60; 50 new vocab; particles を/に/へ/で", "Quiz mode daily until 90%+; handwriting canvas for every new kanji"],
["3", "N5 kanji block 3 + verbs", "SRS: kanji 61–109; 50 vocab (verbs focus); ます-form past/negative, ませんか/ましょう", "Speak: self-introduction until fluent from memory"],
["4", "N5 consolidation", "Full N5 mock quiz; review all SRS backlogs; grammar て-form (てください/ています/てもいいです)", "Take a timed N5 practice test; note weak areas; re-drill only those"],
["5", "Bridge to N4: て-form mastery", "て-form conjugation drills for all verb groups; 50 new N4 vocab; 〜てから/〜まえに", "Write a daily 3-sentence diary in Japanese by hand"],
["6", "Plain form", "Dictionary form, ない form, た form; 〜たことがあります/〜たり〜たり; 50 vocab", "Shadow longer sentences; record and compare with TTS audio"],
["7", "N4 kanji wave 1", "First 50 N4 kanji into SRS; comparisons より/のほうが/いちばん; 50 vocab", "Handwrite kanji compounds, not just single characters"],
["8", "N4 kanji wave 2 + potential", "N4 kanji 51–100; potential form; 〜つもり/〜でしょう; 50 vocab", "Mock JFT-Basic section; workplace dialogue practice (interview Q&A)"],
["9", "N4 kanji wave 3 + giving/receiving", "N4 kanji 101–150; あげる/くれる/もらう; 〜んです; 50 vocab", "Write a short self-introduction letter (rirekisho-style intro)"],
["10", "Listening + occupational Japanese", "Daily listening (NHK Easy, shadowing sets); workplace phrases for your track (SSW/IT)", "Listen-first days: 30 min listening before any reading"],
["11", "Full mock exams", "Complete timed N4 mock (or JFT-Basic mock); review every error into SRS", "Simulate test-day conditions; target 70%+ before proceeding"],
["12", "Polish + apply", "Clear all SRS backlogs; re-take weakest mock sections; speaking interview simulation", "Book the real JLPT/JFT-Basic; prepare requirements (DMW-licensed agency for SSW)"],
];

const doc = new Document({
  styles: { default: { document: { run: { font: LATIN, size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    children: [
      // Cover
      new Paragraph({ spacing: { after: 1600 } }),
      P("TaraGO Study Pack", { size: 56, bold: true, align: AlignmentType.CENTER, color: "C0392B", after: 200 }),
      P("JLPT N5 Review + Road to N4", { size: 32, align: AlignmentType.CENTER, after: 200 }),
      P("Kanji • Vocabulary • Grammar • Speaking • Writing", { size: 24, align: AlignmentType.CENTER, color: "7F8C8D", after: 200 }),
      P("Printable companion to the TaraGO N5 Interactive Reviewer", { size: 22, align: AlignmentType.CENTER, color: "7F8C8D" }),
      pageBreak(),

      // How to use
      H1("How to Use This Pack"),
      P("This pack applies three research-backed rules. 1) Space it out: review a little every day — the schedule below tells you when to revisit material, because memory decays fastest right after learning (Ebbinghaus; Cepeda et al., 2006). 2) Test yourself: always try to recall before looking at the answer — retrieval practice beats re-reading (Roediger & Karpicke, 2006). 3) Produce, don't just recognize: write kanji by hand and say sentences aloud — production builds memories that recognition alone cannot (Longcamp et al., 2005; Swain, 1985)."),
      P("Daily routine (45–60 min): 20 min SRS flashcards (app) → 10 min handwriting (this pack's kanji sheets) → 10 min grammar point of the day → 10 min shadowing aloud → 5 min quiz."),

      H1("12-Week Study Plan: N5 Refresher → N4-Ready"),
      P("Assumes you have already finished an N5 video course. Weeks 1–4 rebuild and consolidate N5; weeks 5–12 bridge to N4, the JLPT level required for the SSW visa (with JFT-Basic as the alternative test). Engineer/IT-track learners: continue the same method toward N3 after week 12.", { italics: true }),
      tbl([700, 2300, 4200, 2880], ["Wk", "Theme", "Study targets", "Practice focus"], PLAN),
      pageBreak(),

      // Kanji
      H1(`N5 Kanji Master List (${KANJI.length})`),
      P("For each kanji: cover the meaning column, recall it, then write the character 5 times by hand while saying a reading aloud.", { italics: true }),
      tbl([900, 2800, 3200, 3180], ["Kanji", "On'yomi", "Kun'yomi", "Meaning"], KANJI.map(k => [k[0], k[1], k[2], k[3]])),
      pageBreak(),

      // Grammar
      H1(`N5 Grammar Summary (${GRAMMAR.length} points)`),
      ...GRAMMAR.flatMap(g => [
        H2(g[0]),
        P(g[1], { after: 60 }),
        P(g[2], { jp: true, after: 40 }),
        P(g[3], { italics: true, color: "7F8C8D", after: 160 }),
      ]),
      pageBreak(),

      // Vocab
      H1(`N5 Core Vocabulary (${VOCAB.length} words)`),
      ...[...new Set(VOCAB.map(v => v[3]))].flatMap(cat => {
        const rows = VOCAB.filter(v => v[3] === cat).map(v => [v[0], v[1], v[2]]);
        return [H2(cat.charAt(0).toUpperCase() + cat.slice(1) + ` (${rows.length})`), tbl([3200, 3400, 3480], ["Word", "Reading", "Meaning"], rows)];
      }),
      pageBreak(),

      // Writing practice
      H1("Writing Practice Guide"),
      P("Stroke-order basics: top before bottom; left before right; horizontal before vertical when they cross; outside before inside; enclosures closed last; center before symmetric wings."),
      P("Weekly writing tasks:", { bold: true, after: 60 }),
      P("• Weeks 1–4: copy each new kanji 5×, then write it from memory 3× the next day (spaced production)."),
      P("• Weeks 5–8: write a 3-sentence daily diary (today's date, one thing you did, one thing you will do)."),
      P("• Weeks 9–12: write your self-introduction (jikoshoukai) paragraph, then employment-form fields: name in katakana, address format, education history lines — the exact writing you will need for Japanese job applications."),
      H1("Speaking Practice Guide"),
      P("Use the Speak tab in the interactive reviewer. Method (Hamada, 2019): 1) listen once, eyes closed; 2) listen and read along; 3) shadow — speak 1 second behind the audio, copying rhythm and pitch, 5 repetitions; 4) say it alone from memory; 5) record with the mic check and compare. Do one sentence set daily; re-do your self-introduction weekly until it is automatic — it is the first 30 seconds of every job interview in Japan."),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync("TaraGO_Study_Pack.docx", buf); console.log("OK", buf.length); });
