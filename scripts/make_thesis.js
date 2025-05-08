const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, PageBreak,
  LevelFormat, TabStopType, ShadingType
} = require("docx");
const fs = require("fs");

const FONT = "Times New Roman";
const SZ = 22; // 11pt

function P(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 160, line: 300 },
    indent: opts.indent,
    children: [new TextRun({ text, font: FONT, size: opts.size || SZ, bold: opts.bold, italics: opts.italics })],
  });
}
function Pruns(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 160, line: 300 },
    children: runs.map(r => new TextRun({ font: FONT, size: SZ, ...r })),
  });
}
function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: "000000" })],
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: "000000" })],
  });
}
function centered(text, opts = {}) {
  return P(text, { ...opts, align: AlignmentType.CENTER });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ---------------- REFERENCES (IEEE, order of first citation) ----------------
const refs = [
  /*1*/ 'Immigration Services Agency of Japan, "How to obtain the Specified Skilled Worker status of residence," Support Website for the Specified Skilled Worker Program. [Online]. Available: https://www.ssw.go.jp/en/about/sswv/',
  /*2*/ 'Ministry of Economy, Trade and Industry (METI), "Survey on IT human resources supply and demand," Tokyo, Japan, 2019.',
  /*3*/ 'The Japan Foundation, "Survey report on Japanese-language education abroad 2021," Tokyo, Japan, 2023. [Online]. Available: https://www.jpf.go.jp/e/project/japanese/survey/result/survey21.html',
  /*4*/ 'R. DLC. Gonzales, "Nihongo no benkyoo: Learning strategies and motivation of Filipino learners of the Japanese language," Layag, vol. 3, no. 1, 1998.',
  /*5*/ 'R. DLC. Gonzales, "Japanese language education in the Philippines: Profile of learners, motivation, issues, and prospects," SSRN Working Paper 2834347, 2016.',
  /*6*/ 'Y. Nagaya, N. Gillin, and D. Smith, "South-East Asian nurses’ experiences under the Economic Partnership Agreement (EPA) in Japan: How language ability affects self-confidence and interpersonal relationships," Journal of Transcultural Nursing, vol. 35, 2024, doi:10.1177/10436596241271133.',
  /*7*/ 'The Japan Foundation, "Japanese-language pre-training program for Indonesian and Filipino candidates for nurses and certified care workers under Economic Partnership Agreements (EPA)." [Online]. Available: https://www.jpf.go.jp/e/project/japanese/education/training/epa/',
  /*8*/ 'H. Ebbinghaus, Memory: A Contribution to Experimental Psychology. New York, NY, USA: Teachers College, Columbia University, 1913 (original work published 1885).',
  /*9*/ 'N. J. Cepeda, H. Pashler, E. Vul, J. T. Wixted, and D. Rohrer, "Distributed practice in verbal recall tasks: A review and quantitative synthesis," Psychological Bulletin, vol. 132, no. 3, pp. 354–380, 2006.',
  /*10*/ 'H. L. Roediger and J. D. Karpicke, "Test-enhanced learning: Taking memory tests improves long-term retention," Psychological Science, vol. 17, no. 3, pp. 249–255, 2006.',
  /*11*/ 'J. D. Karpicke and H. L. Roediger, "Repeated retrieval during learning is the key to long-term retention," Journal of Memory and Language, vol. 57, no. 2, pp. 151–162, 2007.',
  /*12*/ 'J. Burston, "Twenty years of MALL project implementation: A meta-analysis of learning outcomes," ReCALL, vol. 27, no. 1, pp. 4–20, 2015.',
  /*13*/ 'S. Loewen et al., "Mobile-assisted language learning: A Duolingo case study," ReCALL, vol. 31, no. 3, pp. 293–311, 2019.',
  /*14*/ 'R. Vesselinov and J. Grego, "Duolingo effectiveness study: Final report," City University of New York and University of South Carolina, 2012.',
  /*15*/ 'X. Jiang, J. Rollinson, L. Plonsky, E. Gustafson, and B. Pajak, "Evaluating the reading and listening outcomes of beginning-level Duolingo courses," Foreign Language Annals, vol. 54, no. 4, pp. 974–1002, 2021.',
  /*16*/ 'J. Hamari, J. Koivisto, and H. Sarsa, "Does gamification work? — A literature review of empirical studies on gamification," in Proc. 47th Hawaii Int. Conf. on System Sciences (HICSS), Waikoloa, HI, USA, 2014, pp. 3025–3034.',
  /*17*/ 'R. M. Ryan and E. L. Deci, "Self-determination theory and the facilitation of intrinsic motivation, social development, and well-being," American Psychologist, vol. 55, no. 1, pp. 68–78, 2000.',
  /*18*/ 'H. Rose, "L2 learners’ attitudes toward, and use of, mnemonic strategies when learning Japanese kanji," The Modern Language Journal, vol. 97, no. 4, pp. 981–992, 2013.',
  /*19*/ 'G. H. Gamage, "Perceptions of kanji learning strategies: Do they differ among Chinese character and alphabetic background learners?" Australian Review of Applied Linguistics, vol. 26, no. 2, pp. 17–31, 2003.',
  /*20*/ 'M. Longcamp, M.-T. Zerbato-Poudou, and J.-L. Velay, "The influence of writing practice on letter recognition in preschool children: A comparison between handwriting and typing," Acta Psychologica, vol. 119, no. 1, pp. 67–79, 2005.',
  /*21*/ 'M. Naka and H. Naoi, "The effect of repeated writing on memory," Memory & Cognition, vol. 23, no. 2, pp. 201–212, 1995.',
  /*22*/ 'Y. Hamada, "Shadowing: What is it? How to use it. Where will it go?" RELC Journal, vol. 50, no. 3, pp. 386–393, 2019.',
  /*23*/ 'T. T.-N. Ngo, H. H.-J. Chen, and K. K.-W. Lai, "The effectiveness of automatic speech recognition in ESL/EFL pronunciation: A meta-analysis," ReCALL, vol. 36, no. 1, pp. 4–21, 2024.',
  /*24*/ 'M. Swain, "Communicative competence: Some roles of comprehensible input and comprehensible output in its development," in Input in Second Language Acquisition, S. Gass and C. Madden, Eds. Rowley, MA, USA: Newbury House, 1985, pp. 235–253.',
  /*25*/ 'I. S. P. Nation, "The four strands," Innovation in Language Learning and Teaching, vol. 1, no. 1, pp. 2–13, 2007.',
  /*26*/ 'S. D. Krashen, Principles and Practice in Second Language Acquisition. Oxford, U.K.: Pergamon Press, 1982.',
  /*27*/ 'Japanese-Language Proficiency Test (JLPT), "Four key characteristics," Japan Foundation and Japan Educational Exchanges and Services. [Online]. Available: https://www.jlpt.jp/e/about/points.html',
  /*28*/ 'The Japan Foundation, "Japan Foundation Test for Basic Japanese (JFT-Basic)." [Online]. Available: https://www.jpf.go.jp/jft-basic/e/',
  /*29*/ 'Department of Migrant Workers (Republic of the Philippines), "OFW statistics." [Online]. Available: https://dmw.gov.ph/statistics',
  /*30*/ 'Ministry of Health, Labour and Welfare (Japan), "Summary of notifications on the employment of foreign workers," Tokyo, Japan, 2025.',
  /*31*/ 'A. R. Hevner, S. T. March, J. Park, and S. Ram, "Design science in information systems research," MIS Quarterly, vol. 28, no. 1, pp. 75–105, 2004.',
  /*32*/ 'ISO/IEC 25010:2011, "Systems and software engineering — Systems and software Quality Requirements and Evaluation (SQuaRE) — System and software quality models," International Organization for Standardization, 2011.',
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: SZ } } },
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, bottom: 1440, left: 1800, right: 1440 },
      },
    },
    children: [
      // ---------------- TITLE PAGE ----------------
      new Paragraph({ spacing: { after: 2400 } }),
      centered("TaraGO (タラゴー): An Evidence-Based, Multi-Skill Mobile Japanese Language Learning Application for Filipino Fresh Graduates Seeking Employment in Japan", { bold: true, size: 32, after: 600 }),
      centered("A Thesis Proposal Presented to the Faculty of the School of Information Technology", { after: 100 }),
      centered("Mapua University", { after: 100 }),
      centered("In Partial Fulfillment of the Requirements for the Degree of Bachelor of Science in Information Technology", { after: 800 }),
      centered("[AUTHOR NAME]", { bold: true, after: 100 }),
      centered("[STUDENT NUMBER]", { after: 100 }),
      centered("July 2026", { after: 100 }),
      pageBreak(),

      // ---------------- ABSTRACT ----------------
      H1("Abstract"),
      P("The Philippines produces one of the world’s largest pools of employment-motivated Japanese-language learners, yet Filipino fresh graduates aspiring to work in Japan face a fragmented self-study ecosystem: massively popular applications train isolated receptive skills, while the credentialing examinations that gate employment (JLPT and JFT-Basic) assess no speaking or writing at all. This proposal presents TaraGO, a mobile-first Japanese language learning application purpose-built for Filipino fresh graduates targeting three employment pathways: Specified Skilled Worker (SSW) roles requiring JLPT N4/JFT-Basic, engineer and IT positions typically requiring N3–N2, and a general N5-to-N2 progression track. TaraGO integrates five competencies — kanji, vocabulary, grammar, speaking, and writing — in a single spaced-repetition-driven curriculum. Each design decision is grounded in empirical literature: distributed practice and retrieval-based learning for retention; mnemonic and component-analysis strategies for kanji; shadowing with automatic speech recognition feedback for speaking; and stylus/finger handwriting production for orthographic memory. The proposal reviews migration-policy evidence showing that language ability is the decisive determinant of Filipino workers’ occupational success in Japan, analyzes gaps in existing applications, and specifies a design-science methodology with an ISO/IEC 25010-based evaluation plan and a planned quasi-experimental study comparing TaraGO against a conventional single-skill application stack."),
      Pruns([{ text: "Index Terms — ", bold: true, italics: true }, { text: "mobile-assisted language learning, spaced repetition, JLPT, kanji acquisition, automatic speech recognition, shadowing, Filipino migrant workers, Specified Skilled Worker, design science", italics: true }]),
      pageBreak(),

      // ---------------- CHAPTER 1 ----------------
      H1("Chapter 1: Introduction"),
      H2("1.1 Background of the Study"),
      P("Japan has entered a period of structural labor shortage. The Specified Skilled Worker (SSW) residence status, introduced in April 2019, now covers sixteen industrial fields — including nursing care, food service, manufacturing, construction, and agriculture — and admits foreign workers who pass an industry skills test plus a Japanese language requirement of JLPT N4 or the JFT-Basic [1]. In parallel, Japan’s Ministry of Economy, Trade and Industry projects a shortage of up to approximately 790,000 IT professionals by 2030, driving sustained recruitment of foreign engineers, for whom employers typically expect conversational Japanese in the N3–N2 range in addition to technical credentials [2]. Foreign employment in Japan reached a record of roughly 2.3 million workers in 2024–2025 [30]."),
      P("The Philippines is exceptionally positioned to supply this demand. The Japan Foundation’s most recent global survey ranks the Philippines ninth worldwide with 44,457 Japanese-language learners, and Japanese-language education in the country is expanding across universities, TESDA-registered training centers, and secondary schools [3]. Filipino learners are predominantly instrumentally motivated: local studies consistently find that employment in Japan is the primary driver of Japanese language study among Filipino university students [4], [5]. The Department of Migrant Workers records steady deployment of Filipino workers to Japan across healthcare, technical, and skilled trades [29]."),
      P("Yet the pathway from graduation to employment in Japan is bottlenecked by language. Under the Japan–Philippines Economic Partnership Agreement (JPEPA), Filipino nurse candidates historically passed Japan’s national licensure examination at single-digit rates — about 7% in the program’s early cohorts — with insufficient Japanese proficiency identified as the dominant cause, despite six months of language training in the Philippines and six months in Japan [6], [7]. Qualitative research on Southeast Asian nurses under the EPA further shows that limited Japanese undermines self-confidence and workplace relationships even after arrival [6]. Language, in short, is not one requirement among many; it is the gating requirement."),
      H2("1.2 Statement of the Problem"),
      P("Filipino fresh graduates who self-study Japanese face four compounding problems. First, fragmentation: the de facto standard self-study stack splits skills across separate tools (e.g., one application for kanji recognition, another for grammar drills, a third for vocabulary flashcards), none of which sequences content toward the JLPT levels that Japanese employers and immigration authorities actually require. Second, receptive-skill bias: the JLPT and JFT-Basic assess vocabulary, grammar, reading, and listening through selected-response formats and contain no speaking or writing component [27], [28], so neither the tests nor the applications that target them train the productive skills that determine workplace success [6]. Third, retention decay: learners who complete beginner materials — such as an N5 video course — have no structured, evidence-based mechanism for reviewing and consolidating that knowledge before advancing, despite a century of memory research showing rapid forgetting without spaced review [8], [9]. Fourth, context mismatch: global applications are not localized for Filipino learners’ needs — they do not leverage Filipino/English bilingual scaffolding, do not align with SSW/JPEPA/engineer visa requirements, and do not include the occupational Japanese (caregiving, factory, IT, keigo) that Filipino deployment pathways demand."),
      P("Formally: there exists no integrated, evidence-based mobile application that takes a Filipino fresh graduate from post-N5 review to employment-ready Japanese across kanji, vocabulary, grammar, speaking, and writing, aligned with the credentialing requirements of Japan’s employment pathways."),
      H2("1.3 Objectives of the Study"),
      P("General objective: to design and develop TaraGO, a mobile Japanese language learning application for Filipino fresh graduates seeking employment in Japan, and to evaluate its usability and learning effectiveness.", { after: 80 }),
      P("Specific objectives:", { after: 80 }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ font: FONT, size: SZ, text: "To implement a spaced-repetition scheduling engine with active-recall testing for kanji and vocabulary, grounded in distributed-practice and testing-effect research [9], [10], [11];" })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ font: FONT, size: SZ, text: "To implement a structured grammar curriculum sequenced N5 through N2, with contextualized example sentences and micro-lessons;" })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ font: FONT, size: SZ, text: "To implement a speaking module combining shadowing exercises with automatic speech recognition (ASR) feedback [22], [23];" })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ font: FONT, size: SZ, text: "To implement a writing module with stroke-order guidance and handwriting production practice for kana and kanji [20], [21];" })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ font: FONT, size: SZ, text: "To align content tracks with the three dominant Filipino employment pathways (SSW/N4, engineer/N3–N2, general N5→N2) and embed occupational Japanese modules;" })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 160 }, children: [new TextRun({ font: FONT, size: SZ, text: "To evaluate the application using ISO/IEC 25010 quality characteristics and a quasi-experimental pretest/posttest comparison against a conventional multi-application self-study stack [31], [32]." })] }),
      H2("1.4 Significance of the Study"),
      P("For Filipino fresh graduates, TaraGO converts a fragmented, trial-and-error self-study process into a single guided pathway whose milestones correspond to real hiring gates (JFT-Basic/N4 for SSW; N3–N2 for engineering roles). For the Philippine workforce-development ecosystem (TESDA training centers, DMW-licensed sending organizations, university foreign-language departments), the application offers a scalable supplement to classroom instruction whose pre-departure language outcomes directly affect deployment success [7], [29]. For the academic community, the study contributes a documented case of translating cognitive-science and second-language-acquisition findings into concrete mobile-application design decisions, and empirical evidence on whether an integrated five-skill design outperforms the prevailing single-skill application stack. For Japanese employers and Japanese society, better-prepared workers reduce onboarding cost and the well-documented psychosocial burden that language difficulty imposes on migrant workers [6]."),
      H2("1.5 Scope and Limitations"),
      P("The study covers the design, development, and evaluation of a cross-platform mobile application spanning JLPT levels N5 through N2, with the initial content build-out and evaluation focused on the N5 review → N4 progression (the SSW gate). Speaking assessment uses on-device/cloud ASR rather than human raters and therefore measures segmental accuracy and fluency proxies rather than full communicative competence [23]. Writing evaluation covers kana and kanji character production and short guided composition, not free academic writing. The evaluation cohort will be drawn from Filipino college students and fresh graduates studying Japanese; findings may not generalize to other L1 groups. The application supplements rather than replaces formal instruction; claims are limited to self-study effectiveness."),
      pageBreak(),

      // ---------------- CHAPTER 2 ----------------
      H1("Chapter 2: Review of Related Literature"),
      H2("2.1 Labor Migration Context: Why Language Is the Gating Skill"),
      P("Japan’s SSW framework makes the language requirement explicit and testable: JLPT N4 or JFT-Basic is a statutory precondition for the visa [1], [28]. The empirical record of the JPEPA healthcare corridor demonstrates what happens when language preparation is inadequate. Filipino nurse candidates — all licensed professionals in the Philippines — passed the Japanese national board examination at rates as low as 1 of 59 (2010) and 1 of 113 (2011), a failure attributed overwhelmingly to Japanese-language demands rather than clinical knowledge [6], [7]. Nagaya, Gillin, and Smith’s qualitative study of EPA nurses found that language difficulty eroded self-confidence and workplace relationships even among those who cleared entry requirements [6]. Studies of SSW and technical-intern populations similarly identify language barriers as a primary stressor even at intermediate (N3) proficiency. The implication for system design is direct: an application for employment-bound Filipino learners must target not merely test passage but functional, productive proficiency."),
      H2("2.2 Japanese-Language Education in the Philippines"),
      P("The Philippines hosts one of the world’s largest Japanese-learning populations (44,457 learners; ninth globally) [3]. Gonzales’ studies of Filipino learners across Metro Manila universities established two persistent findings: motivation is predominantly instrumental — oriented toward employment and study in Japan — and learners favor direct strategies (memorization, repetition, practical use) [4], [5]. Instrumentally motivated learners are well-served by curricula with visible, career-linked milestones; this motivational profile directly informs TaraGO’s pathway architecture, in which every unit is tagged to a JLPT level and an employment gate. The same literature documents chronic constraints in Philippine Japanese-language education — teacher shortages, limited contact hours, and uneven access outside Metro Manila [5] — conditions under which a well-designed mobile application has the greatest marginal value."),
      H2("2.3 Cognitive Foundations: Spacing, Testing, and the Forgetting Curve"),
      P("Ebbinghaus first quantified the exponential decay of unrehearsed memory [8]. A century of subsequent research has produced two of the most robust effects in cognitive psychology. The spacing effect — distributed practice outperforms massed practice — was confirmed by Cepeda et al.’s quantitative synthesis of 254 studies involving over 14,000 participants [9]. The testing effect — retrieval practice produces stronger long-term retention than restudy — was demonstrated by Roediger and Karpicke, who showed that repeated testing outperformed repeated studying on delayed tests even though studying appeared superior at short delays [10]; Karpicke and Roediger further showed that repeated retrieval, not repeated exposure, is the key driver of retention [11]. These findings jointly prescribe the core mechanic of TaraGO: all kanji and vocabulary knowledge is scheduled through an expanding-interval spaced-repetition system in which every exposure is an active retrieval attempt, not a passive review. This is precisely the mechanism the user-facing problem demands — a learner returning from a completed N5 course needs systematic reactivation of partially decayed knowledge, which spaced retrieval provides more efficiently than re-watching course material [9], [10]."),
      H2("2.4 Mobile-Assisted Language Learning and Gamification"),
      P("Burston’s meta-analysis of twenty years of MALL implementations found positive learning outcomes concentrated in studies with sound pedagogical design, particularly vocabulary interventions using spaced flashcard systems [12]. Duolingo research illustrates both promise and limits of consumer applications: Vesselinov and Grego reported measurable gains per study hour [14], and Jiang et al. found beginning-level courses produced reading and listening outcomes comparable to university semesters [15], but Loewen et al.’s independent semester-long study found low completion and limited productive-skill development, with participants noting the absence of speaking practice and grammar explanation [13]. The lesson is not that mobile learning fails, but that engagement mechanics must be married to complete pedagogy. On engagement, Hamari et al.’s review found gamification generally improves motivation and engagement, with effects contingent on implementation quality and user context [16]. Self-determination theory provides the design compass: gamification should feed competence (visible progress toward meaningful milestones), autonomy (learner-selected pathways), and relatedness (peer cohorts), rather than exploit extrinsic reward loops [17]. TaraGO’s gamification is therefore milestone-based — progress bars toward JFT-Basic/N4/N3/N2 gates and simulated mock-exam readiness scores — rather than streak-anxiety mechanics."),
      H2("2.5 Kanji Acquisition"),
      P("Kanji is the highest-attrition component of Japanese study for learners from alphabetic L1 backgrounds such as Filipino learners. Gamage found that alphabetic-background learners rely more heavily on rote repetition and benefit from explicit strategy instruction compared to Chinese-character-background learners [19]. Rose’s study of mnemonic strategy use showed mnemonics are effective when meaningfully connected to kanji components, but over-elaborate mnemonics can detach meaning from reading [18]. The synthesis adopted for TaraGO: teach kanji through radical/component decomposition with lightweight mnemonics, always paired with readings in high-frequency vocabulary context, and schedule recognition and recall as separate SRS items — an approach consistent with both strategy research and the retrieval-practice literature [10], [18], [19]."),
      H2("2.6 Speaking: Shadowing and Automatic Speech Recognition"),
      P("Because the JLPT tests no speaking [27], employment-bound learners systematically under-train the skill their workplace success most depends on [6]. Two evidence-based, app-deliverable techniques address this. Shadowing — immediate vocalized repetition of auditory input — has a substantial research base, synthesized by Hamada, showing gains in listening comprehension and phonological processing, with growing evidence for pronunciation and fluency benefits [22]. Automatic speech recognition provides the feedback loop self-study otherwise lacks: Ngo, Chen, and Lai’s meta-analysis found a medium overall effect of ASR practice on L2 pronunciation (g = 0.69), with explicit corrective feedback and medium-to-long treatment durations producing the largest gains, strongest for segmental features [23]. Theoretically, both techniques operationalize Swain’s output hypothesis — production forces learners to notice gaps between their interlanguage and the target [24] — and populate the meaning-focused output and fluency-development strands of Nation’s four-strands framework [25], complementing the comprehensible input emphasized by Krashen [26]."),
      H2("2.7 Writing: Handwriting Production and Orthographic Memory"),
      P("Typing-only curricula leave character knowledge shallow. Longcamp, Zerbato-Poudou, and Velay demonstrated that handwriting practice produces stronger character recognition than typing, attributing the advantage to sensorimotor traces formed during production [20]. Naka and Naoi showed that repeated writing improves memory for graphic designs and unfamiliar characters relative to visual inspection alone [21]. For kana and kanji — where stroke order and internal structure carry linguistic information — these findings justify a writing canvas with stroke-order animation, tracing, and freehand production graded by stroke sequence, positioned as a required (not optional) lane of every kanji unit."),
      H2("2.8 Existing Systems and the Gap"),
      P("The prevailing self-study stack among serious learners pairs a kanji-recognition SRS service (e.g., WaniKani), a grammar-drill service (e.g., Bunpro), and a general flashcard tool (e.g., Anki), optionally alongside a consumer application such as Duolingo. Each is effective in its lane, and each embodies part of the evidence base reviewed above — which is precisely why their limitations are instructive. The kanji services teach recognition but not handwriting or speaking; grammar services assume external instruction; Anki’s scheduling is excellent but content quality depends entirely on user-assembled decks and its learning curve deters novices; consumer apps under-serve grammar explanation and productive skills [13]. None sequences content against JLPT employment gates, none includes occupational Japanese for caregiving, manufacturing, or IT contexts, and none is localized for Filipino learners. Table 1 (Chapter 4) maps these gaps against TaraGO’s feature set. The gap this study addresses is therefore not the absence of good tools but the absence of integration: a single application uniting evidence-based mechanics across all five competencies, aligned to the credentialing reality of labor migration from the Philippines to Japan."),
      pageBreak(),

      // ---------------- CHAPTER 3 ----------------
      H1("Chapter 3: Theoretical Framework and Methodology"),
      H2("3.1 Theoretical Framework"),
      P("TaraGO’s design rests on four interlocking theoretical commitments. (1) Memory science: all discrete knowledge (kanji, vocabulary, grammar patterns) is acquired through spaced retrieval practice [9], [10], [11]. (2) Balanced skill development: the curriculum allocates time across Nation’s four strands — meaning-focused input, meaning-focused output, language-focused learning, and fluency development [25] — correcting the receptive bias induced by JLPT-oriented study [27]. (3) Output-driven noticing: speaking and writing tasks are positioned as engines of acquisition, not assessments, following Swain [24]. (4) Self-determination: engagement mechanics target competence, autonomy, and relatedness [17] and are anchored to instrumentally meaningful milestones, matching the documented motivational profile of Filipino learners [4], [5]."),
      H2("3.2 Research Design"),
      P("The study follows the design-science research paradigm of Hevner et al., in which an innovative artifact is built to address an identified organizational/human problem and rigorously evaluated [31]. Development proceeds in agile increments: (Increment 1) N5 review module — SRS engine, kanji/vocabulary/grammar content, quiz engine; (Increment 2) speaking (shadowing + ASR) and writing (stroke canvas) modules; (Increment 3) N4 curriculum and SSW occupational modules; (Increment 4) N3–N2 tracks and IT/keigo modules."),
      H2("3.3 Evaluation Plan"),
      P("Evaluation has two arms. Arm A (software quality): expert and user evaluation against ISO/IEC 25010 characteristics — functional suitability, performance efficiency, usability, reliability, and portability — using standardized instruments and task-based usability testing with at least 30 representative users [32]. Arm B (learning effectiveness): a quasi-experimental pretest/posttest design over eight weeks with Filipino learners who have completed an N5-level course. The treatment group studies with TaraGO; the comparison group uses a conventional self-assembled stack. Outcome measures: JLPT N5/N4 mock examinations (vocabulary, grammar, reading, listening), a kanji production test (handwriting), and an ASR-scored speaking task, administered at baseline, week 4, and week 8, with a delayed retention test at week 12 — the delayed test being essential because spacing and testing advantages emerge most strongly at delay [9], [10]. Analysis: mixed-design ANOVA on gain scores; usability results reported descriptively against ISO/IEC 25010 sub-characteristics."),
      pageBreak(),

      // ---------------- CHAPTER 4 ----------------
      H1("Chapter 4: The Proposed System — TaraGO"),
      H2("4.1 Concept and Name"),
      P("TaraGO fuses Filipino “Tara!” (“let’s go!”) with the Japanese suffix 語 (go, “language”) — as in Nihongo (日本語) and Tagalogo (タガログ語): an invitation, in one word, for Filipinos to go to the language and go to Japan. The application is mobile-first (Android priority, reflecting Philippine device share), offline-capable (content packs downloadable over intermittent connectivity), and bilingual (English/Filipino explanations)."),
      H2("4.2 Learner Pathways"),
      P("On onboarding, the learner selects a pathway; the curriculum engine then sequences identical core content with different emphases and occupational overlays: (a) SSW Pathway — target JFT-Basic/JLPT N4; overlays for caregiving, food service, and manufacturing vocabulary and dialogue; (b) Engineer/IT Pathway — target N3–N2; overlays for IT vocabulary, business email, and keigo; (c) General Pathway — full N5→N2 progression. Every unit displays its position relative to the learner’s employment gate, operationalizing instrumental motivation [4], [17]."),
      H2("4.3 Core Modules"),
      P("Kanji Module. Kanji are introduced by frequency and JLPT level through component decomposition with concise mnemonics [18], [19]. Each kanji generates four SRS item types: meaning recognition, reading recognition (in vocabulary context), audio-to-character matching, and handwritten production on the stroke canvas [20], [21]. The scheduler treats recognition and production as independent memories.", { after: 120 }),
      P("Vocabulary Module. High-frequency, level-tagged vocabulary is taught in example-sentence context and scheduled through the same expanding-interval SRS with active recall in both directions (JP→EN/FIL, EN/FIL→JP) [9], [11]. Occupational overlays inject pathway-specific lexicons.", { after: 120 }),
      P("Grammar Module. Micro-lessons (3–5 minutes) present one pattern each with plain-language explanation, contrastive notes for Filipino/English speakers, and immediate cloze and sentence-construction exercises; patterns then enter the SRS as cloze items. Sequencing follows JLPT levels N5→N2.", { after: 120 }),
      P("Speaking Module. Each unit ships model audio at native and slowed speeds. The learner progresses through listen → synchronized reading → shadowing → independent production [22]. ASR scores segmental accuracy and provides explicit corrective feedback, the configuration with the strongest meta-analytic support [23]. Weekly “workplace dialogue” scenarios (self-introduction, interview answers, caregiving handover, IT stand-up) build employment-relevant fluency.", { after: 120 }),
      P("Writing Module. Stroke-order animation, guided tracing, then freehand production scored on stroke count/sequence; short guided composition tasks (self-introduction paragraph, simple forms, résumé fields) connect character production to employment documents [20], [21].", { after: 120 }),
      P("Assessment and Review. Daily review queue (SRS due items across all modules); weekly mixed quizzes; full JLPT-format mock examinations per level; a readiness dashboard estimating preparedness per employment gate."),
      H2("4.4 Feature–Evidence Matrix"),
      new Table({
        width: { size: 9000, type: WidthType.DXA },
        columnWidths: [3000, 3600, 2400],
        rows: [
          new TableRow({
            children: ["Feature", "Design decision", "Evidence"].map(t => new TableCell({
              width: { size: t === "Design decision" ? 3600 : t === "Feature" ? 3000 : 2400, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: "DDDDDD" },
              children: [P(t, { bold: true, after: 40, align: AlignmentType.LEFT })],
            })),
          }),
          ...[
            ["Expanding-interval SRS", "All discrete knowledge scheduled by spaced retrieval; every exposure is a test", "[8], [9], [10], [11]"],
            ["Kanji component method", "Radical decomposition + light mnemonics, readings taught in vocabulary context", "[18], [19]"],
            ["Handwriting canvas", "Required production lane for kana/kanji with stroke-order scoring", "[20], [21]"],
            ["Shadowing + ASR", "Staged shadowing with explicit ASR corrective feedback", "[22], [23], [24]"],
            ["Four-strand balance", "Time budget across input, output, language focus, fluency", "[25], [26]"],
            ["Milestone gamification", "Progress toward JFT-Basic/N4/N3/N2 gates; no streak-anxiety loops", "[16], [17]"],
            ["Pathway localization", "SSW/engineer/general tracks; FIL/EN scaffolding; occupational lexicons", "[1], [4], [5], [6]"],
          ].map(([a, b, c]) => new TableRow({
            children: [[a, 3000], [b, 3600], [c, 2400]].map(([t, w]) => new TableCell({
              width: { size: w, type: WidthType.DXA },
              children: [P(t, { after: 40, align: AlignmentType.LEFT })],
            })),
          })),
        ],
      }),
      P("Table 1. Mapping of TaraGO features to supporting literature.", { italics: true, align: AlignmentType.CENTER, after: 240 }),
      H2("4.5 Architecture Overview"),
      P("The system comprises a cross-platform mobile client (offline-first, with local SRS state and downloadable content packs), a backend content service (curriculum, occupational overlays, mock exams), an ASR service (on-device recognition where available, cloud fallback), and an analytics service supporting the evaluation plan. Learner data is minimized and stored in compliance with the Philippine Data Privacy Act of 2012."),
      pageBreak(),

      // ---------------- CHAPTER 5 ----------------
      H1("Chapter 5: Conclusion and Recommendations"),
      P("This proposal has argued from migration-policy evidence, Philippine education research, and cognitive and second-language-acquisition science to a single conclusion: Filipino fresh graduates bound for Japanese employment need an integrated, five-skill, credential-aligned learning application, and the components required to build one are individually well-evidenced but nowhere combined. The JPEPA record shows the human cost of inadequate language preparation [6], [7]; the spacing and testing literature supplies the retention engine [9], [10], [11]; kanji strategy research, handwriting research, and the shadowing/ASR literature supply the skill-specific mechanics [18]–[23]; and self-determination theory supplies the engagement architecture matched to Filipino learners’ instrumental motivation [4], [17]."),
      P("Recommendations for future work: (1) longitudinal tracking of TaraGO users through actual JLPT/JFT-Basic attempts and employment outcomes, closing the loop between app metrics and the real-world gates the app targets; (2) partnership pilots with TESDA-registered training centers and DMW-licensed agencies to test blended deployment; (3) extension of the occupational overlay system to additional SSW fields; (4) investigation of AI-driven conversational practice as a complement to scripted shadowing once base fluency is established."),
      pageBreak(),

      // ---------------- REFERENCES ----------------
      H1("References"),
      ...refs.map((r, i) => new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120, line: 276 },
        indent: { left: 720, hanging: 720 },
        children: [new TextRun({ text: `[${i + 1}] ${r}`, font: FONT, size: SZ })],
      })),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("TaraGO_Thesis_Proposal.docx", buf);
  console.log("OK", buf.length);
});
