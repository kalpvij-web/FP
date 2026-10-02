const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, PageBreak,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, Footer, PageNumber,
  NumberFormat, PositionalTab, PositionalTabAlignment, PositionalTabRelativeTo, PositionalTabLeader,
  LevelFormat, TabStopType,
} = require("docx");

const pages = fs.existsSync("pages.json") ? JSON.parse(fs.readFileSync("pages.json")) : {};
const FONT = "Times New Roman";
const SZ = 24; // 12 pt
const LINE = 360; // 1.5 spacing

// inline markup: **bold**, *italic*
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), ...base, bold: true }));
    else out.push(new TextRun({ text: t.slice(1, -1), ...base, italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...base }));
  return out;
}
const P = (t, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: o.after ?? 120, before: o.before ?? 0 },
  indent: o.noIndent ? undefined : { firstLine: 567 },
  keepNext: o.keepNext,
  children: runs(t, o.run || {}),
});
const C = (t, o = {}) => P(t, { ...o, align: AlignmentType.CENTER, noIndent: true });
const L = (t, o = {}) => P(t, { ...o, align: AlignmentType.LEFT, noIndent: true });
const blank = () => new Paragraph({ spacing: { line: LINE }, children: [] });
const brk = () => new Paragraph({ children: [new PageBreak()] });
const H1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER,
  spacing: { line: LINE, after: 240 }, keepNext: true, children: [new TextRun({ text: t })],
});
const CH = (num, title) => [
  new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 0 }, keepNext: true, children: [new TextRun({ text: num })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 240 }, keepNext: true, children: [new TextRun({ text: title, bold: true })] }),
];
const H2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2, alignment: AlignmentType.LEFT,
  spacing: { line: LINE, before: 120, after: 120 }, keepNext: true, children: [new TextRun({ text: t })],
});
const H3 = (t) => new Paragraph({
  alignment: AlignmentType.LEFT, spacing: { line: LINE, before: 60, after: 60 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true })],
});
const bullet = (t) => new Paragraph({
  numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 60 }, children: runs(t),
});
const num = (t, ref = "num") => new Paragraph({
  numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 60 }, children: runs(t),
});
const caption = (t) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 60 }, keepNext: false,
  children: runs(t, { bold: true }),
});
const source = (t) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 160 }, children: runs(t, { italics: true }),
});
const imgPlaceholder = (label) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { line: LINE, before: 120, after: 60 }, keepNext: true,
  border: { top: { style: BorderStyle.DASHED, size: 6, color: "808080", space: 8 }, bottom: { style: BorderStyle.DASHED, size: 6, color: "808080", space: 8 } },
  children: [new TextRun({ text: `[Insert photograph: ${label}]`, color: "808080" })],
});
const chart = (fn) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, keepNext: true,
  children: [new ImageRun({ type: "png", data: fs.readFileSync(fn), transformation: { width: 520, height: 276 } })],
});

// tables
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: r.map((c, j) => new TableCell({
        width: { size: widths[j], type: WidthType.DXA }, borders,
        shading: i === 0 ? { type: ShadingType.CLEAR, fill: "D9D9D9", color: "auto" } : undefined,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 276 }, children: runs(c, i === 0 ? { bold: true, size: 22 } : { size: 22 }) })],
      })),
    })),
  });
}
const tcaption = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, before: 120, after: 60 }, keepNext: true, children: runs(t, { bold: true }) });

// TOC line with dot leader
const tocLine = (t, key, o = {}) => new Paragraph({
  spacing: { line: LINE, after: 0 }, indent: o.indent ? { left: o.indent } : undefined,
  tabStops: [{ type: TabStopType.RIGHT, position: 8306, leader: "dot" }],
  children: [new TextRun({ text: t + (key === "-" ? "" : "\t" + String(pages[key] ?? "00")), bold: !!o.bold })],
});

// ---------------- PRELIMINARY SECTIONS ----------------
const prelim = [];

// Title page
prelim.push(
  C("**A FIELD PROJECT REPORT ON**", { after: 240 }),
  C("**THE UNSERIOUS GENERATION ON THE MOST SERIOUS TOPIC:**", { after: 0 }),
  C("**IRONY, HUMOUR AND GEN Z LANGUAGE AS POLITICAL RESISTANCE IN THE NEET-UG 2026 PROTESTS**", { after: 360 }),
  C("Submitted in partial fulfilment of the requirements for the", { after: 0 }),
  C("Second Year Bachelor of Arts (S.Y.B.A.) — English Literature", { after: 360 }),
  C("Submitted by", { after: 0 }),
  C("**Gatha Kalpana Vijay**", { after: 360 }),
  C("Under the guidance of", { after: 0 }),
  C("**Prof. Dr. Umesh Jagdale**", { after: 0 }),
  C("Department of English", { after: 480 }),
  C("**Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner**", { after: 0 }),
  C("Affiliated to Savitribai Phule Pune University", { after: 360 }),
  C("**Academic Year 2026–27**", { after: 0 }),
  C("Year of Submission: 2026", { after: 0 }),
  brk(),
);

// Certificate of the College
prelim.push(
  H1("CERTIFICATE"),
  C("Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner", { after: 240 }),
  P("This is to certify that **Gatha Kalpana Vijay**, a student of S.Y.B.A. (English Literature), has satisfactorily completed the Field Project entitled **“The Unserious Generation on the Most Serious Topic: Irony, Humour and Gen Z Language as Political Resistance in the NEET-UG 2026 Protests”** under my guidance and supervision during the academic year 2026–27."),
  P("To the best of my knowledge, this work is the original work of the student and has not been submitted earlier to this or any other university or institution for the award of any degree, diploma or certificate."),
  blank(), blank(),
  L("Place: Sangamner", { after: 0 }),
  L("Date: ____ / ____ / 2026", { after: 480 }),
  sigTable(),
  brk(),
);
function sigTable() {
  const w = [3010, 3010, 3010];
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const nb = { top: none, bottom: none, left: none, right: none };
  const cell = (lines, i) => new TableCell({
    width: { size: w[i], type: WidthType.DXA }, borders: nb,
    children: lines.map((l, k) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 276 }, children: [new TextRun({ text: l, bold: k === 1 })] })),
  });
  return new Table({
    width: { size: 9030, type: WidthType.DXA }, columnWidths: w,
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
    rows: [new TableRow({ children: [
      cell(["____________________", "Project Guide", "Prof. Dr. Umesh Jagdale"], 0),
      cell(["____________________", "Head of Department", "Department of English"], 1),
      cell(["____________________", "Principal", "(Seal of the College)"], 2),
    ] })],
  });
}

// Certificate of Organisation
prelim.push(
  H1("CERTIFICATE OF ORGANISATION"),
  P("**Not Applicable.**", { noIndent: true }),
  P("The present Field Project is a desk-based qualitative study of publicly available digital material (protest photographs, posters and memes circulated on Instagram, Pinterest and BBC News). The study was not carried out within, or on behalf of, any host organisation, and therefore no certificate from an organisation is attached."),
  brk(),
);

// Self-Declaration
prelim.push(
  H1("SELF-DECLARATION"),
  P("I, **Gatha Kalpana Vijay**, a student of S.Y.B.A. (English Literature), Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner, hereby declare that the Field Project Report entitled **“The Unserious Generation on the Most Serious Topic: Irony, Humour and Gen Z Language as Political Resistance in the NEET-UG 2026 Protests”** is my own original work, carried out under the guidance of **Prof. Dr. Umesh Jagdale**, Department of English."),
  P("I further declare that this report has not been submitted, either in part or in full, to any other university or institution for the award of any degree, diploma or certificate. All photographs and posters reproduced in this report are the work of their respective creators and have been credited to their original sources; all secondary sources consulted have been duly acknowledged in the References."),
  blank(), blank(),
  L("Place: Sangamner", { after: 0 }),
  L("Date: ____ / ____ / 2026", { after: 480 }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { line: LINE, after: 0 }, children: [new TextRun("____________________")] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { line: LINE, after: 0 }, children: [new TextRun({ text: "Gatha Kalpana Vijay", bold: true })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { line: LINE, after: 0 }, children: [new TextRun("S.Y.B.A. (English Literature)")] }),
  brk(),
);

// Acknowledgements
prelim.push(
  H1("ACKNOWLEDGEMENTS"),
  P("Every project carries the voices of many people, and this one is no exception. I take this opportunity to express my sincere gratitude to all those who made this Field Project possible."),
  P("First and foremost, I am deeply grateful to my guide, **Prof. Dr. Umesh Jagdale**, Department of English, for his constant guidance, patience and encouragement. His suggestions during the progress review helped me sharpen my research question and gave this project its direction."),
  P("I am thankful to the Principal and to the Head of the Department of English, Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner, for providing the opportunity and the academic environment to undertake this Field Project. I also thank all the faculty members of the Department of English for their support."),
  P("I owe a special acknowledgement to the students, protesters, independent journalists and content creators whose photographs, posters and posts form the corpus of this study. Their creativity is the subject of this project, and their work has been credited to its original sources throughout."),
  P("Finally, I thank my family and friends for their unconditional support, their patience with my endless meme-sharing, and for reminding me that even the most serious work can be done with a little humour."),
  blank(),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { line: LINE, after: 0 }, children: [new TextRun({ text: "Gatha Kalpana Vijay", bold: true })] }),
  brk(),
);

// TOC
prelim.push(
  H1("TABLE OF CONTENTS"),
  new Paragraph({ spacing: { line: LINE, after: 120 }, tabStops: [{ type: TabStopType.RIGHT, position: 8306 }], children: [new TextRun({ text: "Contents\tPage No.", bold: true })] }),
  tocLine("A. PRELIMINARY SECTIONS", "-", { bold: true }),
  tocLine("Certificate of the College", "cert", { indent: 360 }),
  tocLine("Certificate of Organisation", "certorg", { indent: 360 }),
  tocLine("Self-Declaration", "decl", { indent: 360 }),
  tocLine("Acknowledgements", "ack", { indent: 360 }),
  tocLine("List of Figures and Tables", "lof", { indent: 360 }),
  tocLine("B. MAIN BODY", "-", { bold: true }),
  tocLine("Chapter 1: Introduction and Research Methodology", "ch1", { indent: 360 }),
  tocLine("Chapter 2: Theoretical Background and Review of Literature", "ch2", { indent: 360 }),
  tocLine("Chapter 3: Profile of the Study Context", "ch3", { indent: 360 }),
  tocLine("Chapter 4: Data Analysis and Interpretation", "ch4", { indent: 360 }),
  tocLine("Chapter 5: Findings, Conclusion and Suggestions", "ch5", { indent: 360 }),
  tocLine("C. CONCLUDING SECTIONS", "-", { bold: true }),
  tocLine("References", "refs", { indent: 360 }),
  tocLine("Appendices", "app", { indent: 360 }),
  tocLine("Progress Report", "prog", { indent: 360 }),
  brk(),
);
// fix: section headers in TOC should not show page number
// List of figures and tables
const FIGS = [
  ["Figure 1", "Quoting the famous dialogue of Spider-Man", "fig1"],
  ["Figure 2", "Wordplay — the institution’s own name as the punchline", "fig2"],
  ["Figure 3", "Comparing the government to quick-commerce delivery apps", "fig3"],
  ["Figure 4", "Reclaiming Gen Z slang — ‘Baddies are not apolitical’", "fig4"],
  ["Figure 5", "‘Error 404’ — borrowing the language of a broken web page", "fig5"],
  ["Figure 6", "Parody of Dora the Explorer’s question to the audience", "fig6"],
  ["Figure 7", "Pun on ‘leak’ — parody of sanitary pad advertisements", "fig7"],
  ["Graph 1", "Corpus items by source platform", "g1"],
  ["Graph 2", "Rhetorical devices in the analysed items", "g2"],
  ["Graph 3", "Primary institutional target of the analysed items", "g3"],
];
const TABS = [
  ["Table 1.1", "Four-dimensional analytical framework", "t11"],
  ["Table 3.1", "Timeline of the NEET-UG 2026 protest movement", "t31"],
  ["Table 4.1", "Composition of the corpus by source platform", "t41"],
  ["Table 4.2", "Coding summary of the analysed items", "t42"],
  ["Table 4.3", "Frequency of rhetorical devices", "t43"],
  ["Table 4.4", "Cultural source domain borrowed by each item", "t44"],
  ["Table 5.1", "Summary of findings against the research objectives", "t51"],
];
prelim.push(H1("LIST OF FIGURES AND TABLES"), H3("List of Figures and Graphs"));
FIGS.forEach(f => prelim.push(tocLine(`${f[0]}: ${f[1]}`, f[2])));
prelim.push(H3("List of Tables"));
TABS.forEach(t => prelim.push(tocLine(`${t[0]}: ${t[1]}`, t[2])));

// ---------------- MAIN BODY ----------------
const body = [];

// CHAPTER 1
body.push(...CH("CHAPTER 1", "INTRODUCTION AND RESEARCH METHODOLOGY"));
body.push(
  H2("1.1 Introduction"),
  P("When the Gen Z of India swarmed the streets of Jantar Mantar, New Delhi, and various other parts of the country — protesting over the NEET-UG paper leak and the repeated failures of the Indian education system — they carried posters quoting Spider-Man while demanding accountability from the government. Their primary focus was the recurring failures and irregularities of the Indian education system, including paper leaks and youth unemployment. They demanded the resignation of the Union Education Minister in order to hold the government accountable, and they demanded compensation and relief for the families of students who had lost their lives by suicide because of examination turmoil. Six student activists from the All India Students’ Association (AISA), along with education activist Sonam Wangchuk, went on an indefinite hunger strike."),
  P("Amidst the protest, Gen Z did not just march — they memed. They showed up unserious while being deeply serious at the same time. Relatable, trending ‘brainrot’ memes, quotes, slang and reels were used to protest and to put forth their demands. Protesters used their creativity to connect the protest with the memes they have consumed all their lives. Unlike the ‘traditional’ protests that came before, Gen Z showed up as it is — full of humour, satire, sarcasm, rage and wisdom. Informal language, satire, dark humour, viral reel dances, brainrot memes and pop-culture references fuelled the protest."),
  P("It has long been understood that protest belongs to formal language, political leaders and mainstream media. Gen Z broke that stereotype, yet again. For the Union government, the opposition and everyone watching, this was completely ‘out of syllabus’. This new-age, humorous, satire-filled language of Gen Z therefore needs to be studied seriously — which is the task of this Field Project."),
  H2("1.2 Statement of the Research Problem"),
  P("Humour is usually seen as the opposite of seriousness. A protest poster that quotes a cartoon or a superhero film risks being dismissed as trivial, childish or ‘unserious’. Yet during the NEET-UG 2026 protests, such posters were among the most widely shared and remembered images of the movement. This creates a puzzle: how can a register that appears unserious carry one of the most serious demands a generation can make — justice in education and accountability from the State? This project addresses that puzzle."),
  H2("1.3 Research Question"),
  P("The central research question of this project is:", { keepNext: true }),
  P("*What made irony and humour an effective register for political resistance in the NEET-UG 2026 protests, rather than undermining the protest’s seriousness?*", { align: AlignmentType.CENTER, noIndent: true }),
  H2("1.4 Thesis Statement"),
  P("This project argues that the memes, reels, slang and language of Gen Z functioned as a **rhetorical mode of political resistance**, not as decoration around the protest. Humour was the method, not the distraction. In short: *the unserious generation on the most serious topic.*"),
  H2("1.5 Objectives of the Study"),
  num("To document the protest language — posters, memes and slang — used by Gen Z during the NEET-UG 2026 protests.", "obj"),
  num("To identify the rhetorical devices (irony, satire, parody, intertextuality, wordplay, slang) used in this protest language.", "obj"),
  num("To identify which institutional failures each item targets.", "obj"),
  num("To examine how the implied audience (in-group and out-group) shapes the meaning of each item.", "obj"),
  num("To explain what humour allowed protesters to say that formal protest language could not, and to help non-Gen Z readers understand this new-age language.", "obj"),
  H2("1.6 Research Design and Approach"),
  P("The study follows a **qualitative, descriptive and interpretive research design**. It is grounded in **multimodal discourse analysis**, an approach that reads images and words together as one combined message. The design is adapted from the three-layer framework of Hakoköngäs, Halmesvaara and Sakki (2020), who analysed the content, form and rhetorical function of internet memes. This project adapts that framework to a very different context — youth protest against institutional failure — and extends it into the four-dimensional close-reading framework shown in Table 1.1."),
  tcaption("Table 1.1: Four-dimensional analytical framework"),
  table([2200, 6830], [
    ["Dimension", "Guiding question applied to each item"],
    ["1. Rhetorical device", "Which device carries the message — irony, satire, parody, intertextuality, wordplay, absurdism, slang reclamation?"],
    ["2. Institutional target", "Which institutional failure or actor does the item accuse?"],
    ["3. Implied audience", "Who is expected to ‘get it’ (in-group), and who is positioned outside the joke (out-group)?"],
    ["4. Function", "What does the item accomplish rhetorically, and what does humour let it say that formal protest language could not?"],
  ]),
  blank(),
  H2("1.7 Sampling Strategy and Sample Size"),
  P("The study uses **purposive (judgement) sampling**. Items were selected for their rhetorical richness, prioritising content that was (a) widely shared, (b) creatively distinct, (c) clearly illustrative of a specific rhetorical device, or (d) highly relatable to Gen Z. Each item is unique, but all of them tie back to the single objective of the movement — accountability and justice."),
  P("The corpus consists of **13 items** (photographs of protest posters, memes and slang) dated between **6 June 2026 and 25 July 2026** — the period from the first physical demonstration at Jantar Mantar to the resignation of the Union Education Minister. Of these, **seven representative items** (Figures 1–7) were selected for detailed close reading in Chapter 4, so that every major rhetorical device in the corpus is covered at least once. The remaining items are documented in Appendix A."),
  H2("1.8 Data Collection Methods"),
  P("The study relies on **secondary data** collected from publicly available digital sources:"),
  bullet("**Instagram** — feeds of protesters, independent journalists and influencers who visited the protest sites (3 items)."),
  bullet("**Pinterest** — boards where protest photographs were collected and re-circulated (9 items)."),
  bullet("**BBC News** — one photograph from news coverage of the protests (1 item)."),
  P("For every item, the original creator or account was recorded so that the item could be credited. A scholarly article (Hakoköngäs et al., 2020) and standard works on memes, humour and resistance were consulted to build the theoretical framework."),
  H2("1.9 Scope and Limitations"),
  bullet("The study is limited to visual and verbal protest material from 6 June to 25 July 2026; it does not cover offline speeches, chants or video reels in depth."),
  bullet("The corpus is small and purposively selected, so the findings are interpretive and cannot be statistically generalised to all protest content."),
  bullet("Because the movement is very recent, little academic scholarship exists on it, and the study relies mainly on primary digital material."),
  bullet("Images on social media are frequently re-posted, so the account credited may be a sharer rather than the original creator."),
  H2("1.10 Chapter Scheme"),
  P("**Chapter 1** introduces the research problem, question, objectives and methodology. **Chapter 2** sets out the theoretical background and reviews the literature. **Chapter 3** profiles the study context — the NEET-UG examination, the National Testing Agency and the 2026 protest movement. **Chapter 4** presents and interprets the data through close readings, tables and graphs. **Chapter 5** presents the findings, conclusion and suggestions."),
  brk(),
);

// CHAPTER 2
body.push(...CH("CHAPTER 2", "THEORETICAL BACKGROUND AND REVIEW OF LITERATURE"));
body.push(
  H2("2.1 Introduction"),
  P("This chapter places the project within existing ideas about internet memes, humour, irony and political resistance. It first explains the key theoretical concepts used in the analysis and then reviews the principal study on which the methodology is built, before identifying the research gap this project fills."),
  H2("2.2 Theoretical Background"),
  H3("2.2.1 Internet Memes as Communication"),
  P("Limor Shifman (2014) defines internet memes as groups of digital items that share common characteristics of content, form and stance, that are created with awareness of each other, and that are circulated, imitated and transformed by many users. Memes are therefore not just jokes; they are a shared language. Ryan M. Milner (2016) describes this as a ‘participatory’ language — anyone who knows the format can join in. This participatory quality is central to understanding why memes became the language of the NEET-UG protests."),
  H3("2.2.2 Multimodality"),
  P("Gunther Kress and Theo van Leeuwen (2006) argue that images, colour, layout and typography carry meaning just as words do. A protest poster is a multimodal text: the handwritten font, the red ‘error’ colour, or the printed image of a cartoon character are all part of its argument. Multimodal discourse analysis reads all these modes together (Kress, 2012)."),
  H3("2.2.3 Intertextuality and Parody"),
  P("Intertextuality, a term introduced by Julia Kristeva, refers to the way every text is built out of other texts. Linda Hutcheon (1985) describes parody as ‘repetition with critical difference’ — a familiar text is repeated, but turned to a new and often critical purpose. When a protester rewrites Spider-Man’s moral lesson or Dora the Explorer’s question, they are using exactly this mechanism."),
  H3("2.2.4 Irony"),
  P("Hutcheon (1994) argues that irony always has an ‘edge’: it says one thing and means another, and it creates communities of those who understand it and those who do not. This idea is used in this project to explain the dual-layered audience of protest humour — an in-group that gets the joke instantly and an out-group that is left outside it."),
  H3("2.2.5 Humour, Carnival and Resistance"),
  P("Mikhail Bakhtin’s (1984) idea of the ‘carnivalesque’ describes how laughter temporarily turns the official world upside down, mocking authority and making the powerful look ridiculous. James C. Scott (1990) shows how subordinate groups use jokes, rumours and disguised speech as ‘weapons’ that are hard for authorities to punish. Majken Jul Sørensen (2016) studies humour in political activism and shows that humour can attract attention, build solidarity among activists and undermine the authority of the powerful. Together, these ideas suggest that humour is not opposed to resistance — it can be one of its most effective forms."),
  H2("2.3 Review of Principal Literature"),
  P("Multimodal discourse analysis provides the primary methodological foundation for this project. Hakoköngäs, Halmesvaara and Sakki (2020), in their article “Persuasion Through Bitter Humor: Multimodal Discourse Analysis of Rhetoric in Internet Memes of Two Far-Right Groups in Finland”, published in *Social Media + Society*, analysed **426 internet memes** posted on Facebook between 2015 and 2017 by two Finnish far-right groups, Finland First and the Soldiers of Odin. They argue that memes function as a distinct mode of political communication — not mere images, but persuasive tools that combine visual and textual elements to construct ideological meaning."),
  P("Their study examined memes along three dimensions: (1) **content** — the thematic categories, which they found to be history, humour, mythology, symbols, news and mottos; (2) **form** — visual composition, colour and font; and (3) **rhetorical function** — the persuasive work each meme performs. They found that irony and humour were deliberately used to crystallise arguments into an easily shareable and concise form, building in-group solidarity while delegitimising opponents, which made memes useful tools for persuasion, mobilisation and attracting new audiences."),
  P("While their study focuses on far-right nationalist rhetoric, this project adapts the same multimodal, three-layer method to a very different context — analysing how Gen Z used irony and humour not to entrench extremist ideology but to sustain a youth-led protest against institutional failure. The ‘bitter humour’ they identify in one context reappears here as what this project calls ‘serious unseriousness’."),
  H2("2.4 Research Gap"),
  P("No existing literature has yet applied this framework to the NEET-UG 2026 Gen Z protests. The movement is recent enough that academic scholarship has not caught up with it. Existing studies of memes and politics tend to focus on Western contexts, on election campaigns, or on extremist groups. Very little attention has been given to Indian student movements or to the specific language of Indian Gen Z — its mixture of English, Hindi, Hinglish, internet slang and global pop culture. This project fills the gap by extending an established method to a new, unstudied case, while helping non-Gen Z generations understand the new-age language of Gen Z."),
  H2("2.5 Conclusion"),
  P("The literature shows that memes are a persuasive, participatory and multimodal form of communication, and that humour has long been used as a weapon of the less powerful. What remains unexamined is how these ideas apply to an Indian, youth-led, education-focused protest. Chapter 3 now describes that context."),
  brk(),
);

// CHAPTER 3
body.push(...CH("CHAPTER 3", "PROFILE OF THE STUDY CONTEXT"));
body.push(
  P("*Note: This Field Project is not attached to a host organisation. In place of an organisational profile, this chapter profiles the institutions and the movement that form the context of the study.*", { noIndent: true }),
  H2("3.1 The NEET-UG Examination"),
  P("The National Eligibility cum Entrance Test (Undergraduate), commonly known as NEET-UG, is the single national entrance examination for admission to undergraduate medical courses in India, including MBBS, BDS and AYUSH programmes. It is one of the largest examinations in the country, taken every year by more than twenty lakh candidates. For many students and families, NEET-UG represents years of preparation, financial sacrifice and emotional pressure, which is why any irregularity in the examination affects the lives of lakhs of young people."),
  H2("3.2 The National Testing Agency (NTA)"),
  P("The National Testing Agency is an autonomous body set up by the Ministry of Education, Government of India, to conduct major entrance examinations, including NEET-UG. As the body responsible for setting, conducting and securing the examination, the NTA — along with the Union Ministry of Education — became the main institutional target of the protests after the paper leak."),
  H2("3.3 The NEET-UG 2026 Paper Leak and the Protest Movement"),
  P("Following the NEET-UG 2026 paper leak, students across India demanded accountability from the government. Their main demands were: (a) accountability for the leak and for the repeated failures of the education system; (b) the resignation of the Union Education Minister; and (c) compensation and relief for the families of students who died by suicide due to examination turmoil. The physical protest was led by the **Cockroach Janta Party (CJP)** — a satirical, meme-based youth collective that offered a humorous counter-narrative to institutional failure — together with various student groups, including the **All India Students’ Association (AISA)**."),
  tcaption("Table 3.1: Timeline of the NEET-UG 2026 protest movement"),
  table([2300, 6730], [
    ["Date", "Event"],
    ["6 June 2026", "Physical demonstrations begin at Jantar Mantar, New Delhi, led by the Cockroach Janta Party (CJP) and various student groups, demanding accountability and the Education Minister’s resignation."],
    ["28 June 2026", "Education activist Sonam Wangchuk joins the Jantar Mantar protest and begins an indefinite hunger strike, alongside six AISA student activists."],
    ["18 July 2026", "Police remove a weakened Sonam Wangchuk from the protest site."],
    ["20 July 2026", "The ‘Sansad Chalo’ (March to Parliament) takes place in New Delhi, resulting in violent clashes with security forces and allegations of brutality against students."],
    ["23 July 2026", "The protests expand into a nationwide agitation across multiple states, including Tamil Nadu and Kerala."],
    ["25 July 2026", "Union Education Minister Dharmendra Pradhan resigns from his position."],
  ]),
  blank(),
  H2("3.4 Gen Z as the Protesting Generation"),
  P("Generation Z — broadly, those born between the late 1990s and the early 2010s — is the first generation to grow up entirely online. Its everyday language is shaped by Instagram reels, Pinterest boards, memes, Marvel films, Hindi-dubbed cartoons, quick-commerce apps and internet slang such as ‘baddie’, ‘brainrot’ and ‘404’. NEET-UG aspirants belong almost entirely to this generation. It was therefore natural that when they protested, they protested in the language they know best — and that this language travelled as quickly online as it did on the streets."),
  H2("3.5 Relevance of the Context to the Study"),
  P("The NEET-UG 2026 protests offer an ideal case for this study because (a) the demands were extremely serious — education, livelihoods and lives; (b) the protesters were overwhelmingly Gen Z; and (c) the movement produced an unusually rich body of humorous, ironic and meme-based protest material. The contrast between the seriousness of the issue and the ‘unseriousness’ of the language is the very contrast this project investigates."),
  brk(),
);

// CHAPTER 4
body.push(...CH("CHAPTER 4", "DATA ANALYSIS AND INTERPRETATION"));
body.push(
  H2("4.1 Introduction"),
  P("This chapter presents the data in three stages. First, it describes the composition of the corpus (Section 4.2). Second, it offers a close reading of seven representative items using the four-dimensional framework from Chapter 1 (Section 4.3). Third, it summarises the coded data in tables and graphs and interprets the patterns that emerge (Sections 4.4 and 4.5)."),
  H2("4.2 Composition of the Corpus"),
  tcaption("Table 4.1: Composition of the corpus by source platform"),
  table([3010, 3010, 3010], [
    ["Source platform", "Number of items", "Percentage of corpus"],
    ["Pinterest", "9", "69.2%"],
    ["Instagram", "3", "23.1%"],
    ["BBC News", "1", "7.7%"],
    ["**Total**", "**13**", "**100%**"],
  ]),
  chart("g1.png"),
  caption("Graph 1: Corpus items by source platform"),
  P("**Interpretation:** More than two-thirds of the corpus was found on Pinterest. This is significant: Pinterest is not a news platform but an aesthetic, collecting platform. The fact that protest posters were saved and re-circulated there alongside fashion and mood boards shows that Gen Z treated protest language as something worth collecting and sharing — protest became part of everyday digital culture rather than something separate from it.", { before: 120 }),
  H2("4.3 Close Reading of Representative Items"),
);

const items = [
  {
    fig: "Figure 1: Quoting the famous dialogue of Spider-Man from the Marvel Universe",
    src: "Source: Instagram @p9xrider and @p9x_puneet_saxena",
    surface: "A protester wearing a Spider-Man mask holds a handwritten poster that reads: *‘With great power comes great responsibility. Can’t say the same for the Indian Government.’*",
    device: "**Intertextuality, irony and parody.** The poster borrows a globally recognised line from the Marvel / Spider-Man franchise. It works through irony: the original line is an earnest moral lesson about the responsibility that comes with power; here it is flipped into an accusation. It is also parody, since the government is cast as a superhero who has the power but has failed to live up to the responsibility that power demands.",
    target: "The government’s failure to act responsibly despite holding power over education policy — especially the mishandling and negligence behind the NEET-UG paper leak. The poster does not argue policy details; it makes a moral accusation using a maxim everyone already knows.",
    audience: "Dual-layered. The **in-group** — Gen Z, raised on Marvel films — understands the reference instantly, without any explanation, which builds cultural fluency and solidarity among protesters. The **out-group** — government officials, older generations and mainstream media — may not register the reference in the same way, which itself marks a generational and cultural gap between the protesters and the institutions they are criticising.",
    why: "The poster compresses a complex critique — a government with power but without accountability — into a single, instantly legible, relatable and shareable line. Formal protest language (slogans, official statements, petitions) has to stay measured and defensible; it cannot openly mock authority without risking being dismissed as disrespectful or ‘illegitimate’. A meme sidesteps this. By routing the accusation through Spider-Man, the protester makes the same claim — *the government has failed its basic responsibility* — but wrapped in something that is ‘just a joke’. This makes it harder to censor, easier to share, relatable to almost every Gen Z viewer and lower-risk for the person posting it. Humour also lowers the barrier to participation: you do not need policy expertise to make or share a meme, so outrage becomes something anyone can contribute to instantly. Formal language can state a grievance, but it cannot spread virally, build belonging through shared laughter, and be simultaneously deniable and cutting.",
  },
  {
    fig: "Figure 2: Wordplay — the institution’s own name as the punchline",
    src: "Source: Instagram @chief.of.shitposting",
    surface: "A protester holds a handwritten poster that reads: *‘The 2 things missing in Education System is EDUCATION & SYSTEM.’*",
    device: "**Wordplay and structural irony.** The poster is built like a riddle: it sets up a question (what is missing?) and delivers a punchline that is simply the name of the institution itself. The very words that are supposed to describe the institution become the accusation against it.",
    target: "The total failure of the education system — not one paper leak, but a system that has become hollow at its core, where neither real education nor a functioning system remains.",
    audience: "Dual-layered. The **in-group** instantly recognises the ‘two things’ riddle format from memes and reels, and the punchline lands without explanation. The **out-group** cannot easily dismiss it, because the joke is not an exaggeration — it is literally the complaint.",
    why: "The poster compresses a sweeping critique of an entire institution into a single, self-evident line. Formal protest language would need statistics, reports and long arguments to claim that the education system has collapsed; the joke does it in eleven words. By turning the institution’s own name against it, the poster makes the accusation almost impossible to argue with — while still being ‘just a joke’.",
  },
  {
    fig: "Figure 3: Comparing the government to quick-commerce delivery apps",
    src: "Source: Pinterest @yaashnaaaa",
    surface: "A black placard held up in the crowd reads: *‘Even Blinkit & Flipkart are faster then your answers.’*",
    device: "**Comparative satire combined with intertextuality from consumer culture.** Blinkit (ten-minute grocery delivery) and Flipkart are apps Gen Z uses every day. The government is compared, unfavourably, to a delivery app.",
    target: "The delay and silence of the government and the National Testing Agency in answering students after the NEET-UG paper leak — while lakhs of students waited for clarity, no answers came.",
    audience: "The **in-group** is urban Gen Z, who live on quick-commerce apps and instantly understand the comparison. For the **out-group**, the comparison exposes how slow and distant the institution looks to the generation it is meant to serve.",
    why: "The poster measures State accountability against the speed Gen Z expects from everyday life, and in that comparison official delay begins to look absurd. Formal language could only say ‘the government has not responded in time’. The meme makes that delay feel ridiculous: a grocery app reaches your door in ten minutes while a national examination body cannot answer in weeks. The absurdity itself becomes the argument.",
  },
  {
    fig: "Figure 4: Reclaiming Gen Z slang — ‘Baddies are not apolitical’",
    src: "Source: Pinterest @Lyaaa",
    surface: "A protester holds up a handwritten poster that reads: *‘BADDIES ARE NOT APOLITICAL.’*",
    device: "**Slang reclamation and subversion of a stereotype.** ‘Baddie’ is Gen Z slang for a confident, stylish young woman — a word usually associated with selfies, fashion and Instagram aesthetics rather than politics.",
    target: "The assumption that Gen Z, and especially young women, are unserious, apolitical and interested only in social media — an assumption institutions rely on in order to ignore youth anger.",
    audience: "The **in-group** proudly self-identifies with the label. The **out-group** — elders, media and politicians — is forced to revise its stereotype.",
    why: "The poster is an identity declaration: it takes the very label used to dismiss Gen Z as ‘unserious’ and turns it into a political stance. In many ways it is the thesis of this project written on a poster — *the unserious generation on the most serious topic*. Formal protest language would have said ‘young women are politically aware’. The slang version says it in the language of the people it describes, which makes it both a statement and a form of belonging: being a ‘baddie’ and being political are no longer opposites.",
  },
  {
    fig: "Figure 5: ‘Error 404’ — borrowing the language of a broken web page",
    src: "Source: Pinterest @FAGUN",
    surface: "A protester holds a poster that reads *‘error 404’* in red, followed by *‘accountability not found.’*",
    device: "**Internet / technological intertextuality and deadpan irony.** The HTTP 404 error is the page every internet user sees when a link is broken or the page they are looking for does not exist. The poster does not shout; it simply ‘reports an error’.",
    target: "The absence of accountability — no minister, no agency and no official taking responsibility for the NEET-UG paper leak.",
    audience: "For the digitally native **in-group**, ‘404’ is read instantly. For the **out-group**, the reference may not register at all — which once again marks the generational and cultural gap between protesters and the institutions they are criticising.",
    why: "The poster frames the government as a broken website: students are searching for accountability, and the page simply does not exist. Formal protest language would say ‘the authorities have failed to take responsibility’. The 404 joke turns that into something visual, compact and instantly shareable, using the red ‘error’ colour itself as part of the message — a clear example of multimodal meaning.",
  },
  {
    fig: "Figure 6: Parody of Dora the Explorer’s question to the audience",
    src: "Source: Pinterest @mydearnikes",
    surface: "A protester holds a printed poster of Dora the Explorer with the line: *‘Kya aapko kahi accountability dikh rahi hai?’* (Do you see accountability anywhere?)",
    device: "**Parody, childhood nostalgia and Hinglish.** The poster parodies Dora’s famous question to the audience from the Hindi-dubbed version of the cartoon.",
    target: "Invisible accountability — the government’s evasion after the NEET-UG leak.",
    audience: "Gen Z viewers who grew up watching Hindi-dubbed Dora on television. The call-and-response format pulls the viewer into answering the question, just as the cartoon did.",
    why: "The poster turns the audience into participants. The obvious answer — ‘No!’ — is the protest itself, and it treats the government’s evasions like a children’s game in which even a child can see what is missing. Formal language can state a demand, but it cannot make the audience say the answer out loud. By borrowing a cartoon every Gen Z viewer remembers, the poster turns a political question into a shared memory, and the shared memory into solidarity.",
  },
  {
    fig: "Figure 7: Pun on ‘leak’ — parody of sanitary pad advertisements",
    src: "Source: Pinterest @Spring potato",
    surface: "A protester in the crowd holds a handwritten poster that reads: *‘MY PAD GOT BETTER LEAK PROTECTION THAN EDUCATION SYSTEM.’*",
    device: "**Pun, parody and dark humour.** The poster puns on the word ‘leak’ — parodying sanitary pad advertisements that promise ‘leak protection’ and placing them beside the NEET-UG paper leak. It also breaks the taboo around menstruation by bringing it into a public protest.",
    target: "The paper leak itself and the complete failure of examination security.",
    audience: "The **in-group**, especially young women, gets the joke instantly. The **out-group** is jolted — both by the comparison and by a taboo subject being made public.",
    why: "Shock combined with humour makes the poster unforgettable: an everyday product outperforms a national system. Formal protest language could never make this comparison; it would be considered ‘improper’. That is exactly why it works — the poster says the education system is so broken that a sanitary pad is more reliable, and it says it in a way no one can unsee.",
  },
];
items.forEach((it, i) => {
  body.push(
    H3(`4.3.${i + 1} ${it.fig.split(": ")[1]}`),
    imgPlaceholder(it.fig.split(":")[0]),
    caption(it.fig),
    source(it.src),
    P(`**Surface content:** ${it.surface}`, { noIndent: true }),
    P(`**Rhetorical device:** ${it.device}`, { noIndent: true }),
    P(`**Institutional target:** ${it.target}`, { noIndent: true }),
    P(`**Implied audience:** ${it.audience}`, { noIndent: true }),
    P(`**Function — why it worked:** ${it.why}`, { noIndent: true, after: 240 }),
  );
});

body.push(
  H2("4.4 Tabulation of the Coded Data"),
  tcaption("Table 4.2: Coding summary of the analysed items"),
  table([800, 1900, 2000, 2230, 2100], [
    ["Fig.", "Item", "Rhetorical device", "Institutional target", "Implied audience"],
    ["1", "Spider-Man quote", "Intertextuality, irony, parody", "Power without accountability", "Marvel-raised Gen Z vs. officials"],
    ["2", "‘Education & System’", "Wordplay, structural irony", "Systemic collapse of education", "Meme-literate Gen Z vs. elders"],
    ["3", "Blinkit & Flipkart", "Comparative satire, intertextuality", "Delay and silence of Govt./NTA", "Urban app-using Gen Z"],
    ["4", "‘Baddies are not apolitical’", "Slang reclamation", "Stereotype of apolitical youth", "Young women; media and elders"],
    ["5", "‘Error 404’", "Tech intertextuality, deadpan irony", "Absence of accountability", "Digital natives vs. out-group"],
    ["6", "Dora the Explorer", "Parody, nostalgia, Hinglish", "Invisible accountability", "Gen Z raised on Hindi Dora"],
    ["7", "Pad ‘leak protection’", "Pun, parody, dark humour", "Paper leak / exam security", "Young women; jolted out-group"],
  ]),
  blank(),
  tcaption("Table 4.3: Frequency of rhetorical devices"),
  table([3500, 2300, 3230], [
    ["Rhetorical device", "Number of items (n = 7)", "Figures"],
    ["Intertextuality", "4", "1, 3, 5, 6"],
    ["Irony", "3", "1, 2, 5"],
    ["Parody", "3", "1, 6, 7"],
    ["Wordplay / pun", "2", "2, 7"],
    ["Comparative satire", "1", "3"],
    ["Slang reclamation", "1", "4"],
    ["Dark humour", "1", "7"],
  ]),
  P("*Note: Most items use more than one device, so the totals exceed the number of items.*", { noIndent: true, align: AlignmentType.LEFT }),
  chart("g2.png"),
  caption("Graph 2: Rhetorical devices in the analysed items"),
  P("**Interpretation:** Intertextuality is the most frequent device, appearing in four of the seven items. Irony and parody follow closely. This shows that the protest humour depended heavily on **borrowing** — Gen Z did not invent new symbols so much as repurpose the culture it already shared (films, cartoons, apps, the internet itself). The borrowed material carried the critique, and the shared recognition of that material created the audience.", { before: 120 }),
  chart("g3.png"),
  caption("Graph 3: Primary institutional target of the analysed items"),
  P("**Interpretation:** Although the jokes are very different on the surface, three of the seven items target the same thing — the absence of accountability — and every other item targets a closely related failure (systemic collapse, official delay, exam security, or the dismissal of youth voices). The humour is varied; the target is single. This is the strongest evidence that the jokes were not random entertainment but a coordinated rhetoric serving one demand.", { before: 120 }),
  tcaption("Table 4.4: Cultural source domain borrowed by each item"),
  table([1500, 3500, 4030], [
    ["Figure", "Cultural source domain", "What it borrows"],
    ["1", "Hollywood superhero film", "A moral maxim about power and responsibility"],
    ["2", "Meme / riddle format", "The set-up and punchline structure"],
    ["3", "Quick-commerce apps", "A benchmark of speed and responsiveness"],
    ["4", "Gen Z internet slang", "An identity label (‘baddie’)"],
    ["5", "Internet / web technology", "The ‘page not found’ error message"],
    ["6", "Children’s television (Hindi dub)", "A call-and-response question to the viewer"],
    ["7", "Television advertising", "The promise of ‘leak protection’"],
  ]),
  blank(),
  H2("4.5 Emerging Patterns"),
  P("Across the seven items analysed, four patterns appear:"),
  num("**Borrowed formats.** Marvel, Dora, delivery apps, 404 pages and pad advertisements — familiar culture carries the critique. The more familiar the source, the faster the message travels.", "pat"),
  num("**Comparison as accusation.** Blinkit and a sanitary pad outperform the State. When everyday products work better than a national institution, the absurdity itself becomes the argument.", "pat"),
  num("**Reclaimed identity.** ‘Baddies are not apolitical’ turns the ‘unserious’ label into a political one; the stereotype used to dismiss Gen Z becomes its banner.", "pat"),
  num("**A single shared target.** Whatever the joke, every item points back to one demand — accountability.", "pat"),
  H2("4.6 Humour as Shield, Sword and Community"),
  P("The analysis suggests that humour performed three simultaneous functions in the protest:"),
  bullet("**Shield** — humour is deniable (‘it’s just a joke’) while remaining unmistakably serious in intent. This protected protesters and made their messages harder to censor."),
  bullet("**Sword** — humour made accusations that formal protest language could not safely articulate, such as comparing the State to a sanitary pad or a broken web page."),
  bullet("**Community** — the moment of ‘you get it’ is a form of political solidarity. Shared laughter turned isolated anger into collective belonging, and lowered the barrier for anyone to participate."),
  brk(),
);

// CHAPTER 5
body.push(...CH("CHAPTER 5", "FINDINGS, CONCLUSION AND SUGGESTIONS"));
body.push(
  H2("5.1 Results and Findings"),
  num("**Humour was the method, not the distraction.** Every analysed item used humour to deliver a serious accusation. None of the items was ‘just a joke’; each had a clear institutional target.", "find"),
  num("**Intertextuality was the dominant device.** Four of seven items borrowed from shared popular culture (Marvel, Dora, quick-commerce apps, the internet), followed by irony and parody (three items each).", "find"),
  num("**The target was single and consistent.** Despite varied surfaces, all items pointed back to one demand: accountability from the government and the National Testing Agency.", "find"),
  num("**Humour created a dual-layered audience.** Each item built instant solidarity within Gen Z (the in-group) while exposing the cultural and generational distance between protesters and institutions (the out-group).", "find"),
  num("**Humour said what formal language could not.** It made accusations deniable, compact, shareable and low-risk, and it allowed comparisons (a pad, a delivery app, a broken web page) that would be considered ‘improper’ in formal protest language.", "find"),
  num("**Humour lowered the barrier to participation.** No policy expertise was needed to make or share a meme, so outrage became something anyone could contribute to.", "find"),
  num("**Protest became part of everyday digital culture.** Over two-thirds of the corpus circulated on Pinterest, showing that protest posters were collected and shared like any other aesthetic content.", "find"),
  num("**Gen Z reclaimed its own image.** Items such as ‘Baddies are not apolitical’ turned the stereotype of the ‘unserious’ generation into a political identity.", "find"),
  tcaption("Table 5.1: Summary of findings against the research objectives"),
  table([3800, 5230], [
    ["Objective", "Finding"],
    ["1. Document Gen Z protest language", "13 items documented from Instagram, Pinterest and BBC News (6 June – 25 July 2026)."],
    ["2. Identify rhetorical devices", "Intertextuality (4), irony (3), parody (3), wordplay/pun (2), satire, slang reclamation and dark humour (1 each)."],
    ["3. Identify institutional targets", "All items target accountability-related failures of the government / NTA."],
    ["4. Examine implied audience", "Consistently dual-layered: in-group solidarity and out-group exclusion."],
    ["5. Explain what humour enabled", "Deniability, compression, shareability, participation and ‘improper’ comparisons."],
  ]),
  blank(),
  H2("5.2 Conclusion"),
  P("This project set out to answer one question: *what made irony and humour an effective register for political resistance in the NEET-UG 2026 protests, rather than undermining the protest’s seriousness?*"),
  P("The answer is that humour did not reduce the seriousness of the protest — it **carried** it. Irony allowed protesters to say one thing and mean another, inviting the in-group to complete the meaning. Parody and intertextuality borrowed the authority of familiar culture and turned it against the State. Wordplay and comparison made institutional failure look absurd, and absurdity is difficult to defend. Slang reclamation converted the stereotype of an ‘unserious’ generation into a political identity. Together, these devices made the protest language compact, shareable, deniable and participatory — qualities that formal protest language structurally lacks."),
  P("In this sense, each meme served as a tiny revolution. The memes, reels, slang and language of Gen Z were not decoration around the protest; they were a rhetorical mode of political resistance in their own right. The movement that began at Jantar Mantar on 6 June 2026 and spread across the country ended with the resignation of the Union Education Minister on 25 July 2026. This study does not claim that memes alone caused that outcome, but it shows that humour was central to how the movement spoke, spread and held together."),
  P("The NEET-UG 2026 protests therefore demonstrate that Gen Z has developed its own grammar of dissent — one that is ironic, multimodal and deeply serious underneath. It is the unserious generation on the most serious topic, and it deserves to be read seriously."),
  H2("5.3 Recommendations and Suggestions"),
  H3("For institutions and policy makers"),
  bullet("Treat humorous and meme-based protest as legitimate political expression, not as trivial or disrespectful content; the jokes contain precise grievances."),
  bullet("Strengthen examination security and transparency, and communicate quickly and clearly with students during a crisis — the ‘Blinkit’ and ‘404’ posters show that silence itself becomes a target."),
  H3("For media and older generations"),
  bullet("Engage with the language of Gen Z rather than dismissing it; understanding the references is the first step to understanding the demands."),
  H3("For educators and the academic community"),
  bullet("Include internet memes, protest posters and digital rhetoric in the study of English language and literature, as they are a living form of satire and multimodal text."),
  bullet("Use multimodal discourse analysis as a classroom tool to develop students’ critical media literacy."),
  H3("Suggestions for further research"),
  bullet("Extend the analysis to video reels, chants, speeches and the Cockroach Janta Party’s online content."),
  bullet("Use a larger corpus and quantitative content analysis to test whether the patterns found here hold across the whole movement."),
  bullet("Compare the NEET-UG 2026 protests with Gen Z protests in other countries to identify a shared global ‘grammar’ of humorous dissent."),
  bullet("Conduct interviews or surveys with protesters to understand the intentions behind the posters."),
  brk(),
);

// ---------------- CONCLUDING SECTIONS ----------------
const end = [];
const ref = (t) => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: LINE, after: 120 }, indent: { left: 567, hanging: 567 }, children: runs(t) });
end.push(
  H1("REFERENCES"),
  P("*(References follow the APA 7th edition citation style.)*", { noIndent: true, align: AlignmentType.LEFT }),
  H3("Books and Journal Articles"),
  ref("Bakhtin, M. (1984). *Rabelais and his world* (H. Iswolsky, Trans.). Indiana University Press. (Original work published 1965)"),
  ref("Hakoköngäs, E., Halmesvaara, O., & Sakki, I. (2020). Persuasion through bitter humor: Multimodal discourse analysis of rhetoric in internet memes of two far-right groups in Finland. *Social Media + Society, 6*(2), 1–11. https://doi.org/10.1177/2056305120921575"),
  ref("Hutcheon, L. (1985). *A theory of parody: The teachings of twentieth-century art forms*. Methuen."),
  ref("Hutcheon, L. (1994). *Irony’s edge: The theory and politics of irony*. Routledge."),
  ref("Kress, G. (2012). Multimodal discourse analysis. In J. P. Gee & M. Handford (Eds.), *The Routledge handbook of discourse analysis* (pp. 35–50). Routledge."),
  ref("Kress, G., & van Leeuwen, T. (2006). *Reading images: The grammar of visual design* (2nd ed.). Routledge."),
  ref("Kristeva, J. (1980). *Desire in language: A semiotic approach to literature and art* (L. S. Roudiez, Ed.). Columbia University Press."),
  ref("Milner, R. M. (2016). *The world made meme: Public conversations and participatory media*. MIT Press."),
  ref("Scott, J. C. (1990). *Domination and the arts of resistance: Hidden transcripts*. Yale University Press."),
  ref("Shifman, L. (2014). *Memes in digital culture*. MIT Press."),
  ref("Sørensen, M. J. (2016). *Humour in political activism: Creative nonviolent resistance*. Palgrave Macmillan."),
  H3("Primary Sources (Corpus)"),
  ref("BBC News. (2026). [Photograph from coverage of the NEET-UG 2026 student protests]. https://www.bbc.com/news/articles/cvg9w9dxd7lo"),
  ref("@chief.of.shitposting. (2026). [Photograph of protest poster]. Instagram. https://www.instagram.com/chief.of.shitposting/"),
  ref("@p9xrider & @p9x_puneet_saxena. (2026). [Photograph of protest poster]. Instagram. https://www.instagram.com/p9xrider/"),
  ref("@purva.thakaree. (2026). [Photograph of protest poster]. Instagram. https://www.instagram.com/purva.thakaree/"),
  ref("Dhruv [@ruvvvdump]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/ruvvvdump/"),
  ref("FAGUN [@epic_kid_]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/epic_kid_/"),
  ref("Lyaaa [@aalyasharm0]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/aalyasharm0/"),
  ref("Miss.0 [@ifathima542]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/ifathima542/"),
  ref("mydearnikes. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/mydearnikes/"),
  ref("prwmmm.era [@inpremmm12]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/inpremmm12/"),
  ref("Sania [@saniaazafar]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/saniaazafar/"),
  ref("Spring potato [@thisisspringpotato]. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/thisisspringpotato/"),
  ref("yaashnaaaa. (2026). [Photograph of protest poster]. Pinterest. https://in.pinterest.com/yaashnaaaa/"),
  brk(),

  H1("APPENDICES"),
  H2("Appendix A: Complete Corpus Register"),
  P("The table below lists all 13 items in the corpus with their sources. Items 1–7 are analysed in detail in Chapter 4; the remaining items were collected for the corpus and are reproduced in Appendix B.", { noIndent: true }),
  table([700, 1500, 3400, 3430], [
    ["No.", "Platform", "Source / credit", "Treatment in report"],
    ["1", "Instagram", "@p9xrider and @p9x_puneet_saxena", "Analysed — Figure 1"],
    ["2", "Instagram", "@chief.of.shitposting", "Analysed — Figure 2"],
    ["3", "Pinterest", "@yaashnaaaa", "Analysed — Figure 3"],
    ["4", "Pinterest", "@Lyaaa (aalyasharm0)", "Analysed — Figure 4"],
    ["5", "Pinterest", "@FAGUN (epic_kid_)", "Analysed — Figure 5"],
    ["6", "Pinterest", "@mydearnikes", "Analysed — Figure 6"],
    ["7", "Pinterest", "@Spring potato (thisisspringpotato)", "Analysed — Figure 7"],
    ["8", "Instagram", "@purva.thakaree", "Corpus — Appendix B"],
    ["9", "Pinterest", "@Miss.0 (ifathima542)", "Corpus — Appendix B"],
    ["10", "Pinterest", "@Dhruv (ruvvvdump)", "Corpus — Appendix B"],
    ["11", "Pinterest", "@Sania (saniaazafar)", "Corpus — Appendix B"],
    ["12", "BBC News", "bbc.com/news/articles/cvg9w9dxd7lo", "Corpus — Appendix B"],
    ["13", "Pinterest", "@prwmmm.era (inpremmm12)", "Corpus — Appendix B"],
  ]),
  blank(),
  H2("Appendix B: Additional Corpus Items"),
  P("Photographs of corpus items 8–13, reproduced with credit to their original sources.", { noIndent: true }),
  ...[["Item 8", "Instagram @purva.thakaree"], ["Item 9", "Pinterest @Miss.0"], ["Item 10", "Pinterest @Dhruv"], ["Item 11", "Pinterest @Sania"], ["Item 12", "BBC News (bbc.com/news/articles/cvg9w9dxd7lo)"], ["Item 13", "Pinterest @prwmmm.era"]]
    .flatMap(([a, b]) => [imgPlaceholder(a), source(`Courtesy: ${b}`)]),
  H2("Appendix C: Close-Reading Instrument"),
  P("The following set of questions was applied to every item in the corpus:", { noIndent: true }),
  num("**Surface content:** What is literally on the poster, meme or reel (text, image, colour, layout)?", "inst"),
  num("**Rhetorical device:** Which device carries the message — satire, irony, parody, absurdism, intertextuality, wordplay, in-group reference, slang?", "inst"),
  num("**Institutional target:** Which institutional failure or actor does it accuse?", "inst"),
  num("**Implied audience:** Who is expected to understand it (in-group), and who is positioned outside it (out-group)?", "inst"),
  num("**Function:** Why did it work — what did the humour let it say that formal protest language could not?", "inst"),
  H2("Appendix D: Glossary of Gen Z Terms Used in the Report"),
  table([2400, 6630], [
    ["Term", "Meaning"],
    ["Baddie", "A confident, stylish young woman; usually associated with fashion and social-media aesthetics."],
    ["Brainrot", "Absurd, repetitive, hyper-online meme content; also the feeling of having consumed too much of it."],
    ["Error 404", "The ‘page not found’ message on the internet; used to say something does not exist."],
    ["Meme", "An image, text or video format that is copied, adapted and shared by many users online."],
    ["Reel", "A short vertical video on Instagram, often set to trending audio."],
    ["Shitposting", "Deliberately absurd or low-effort humorous posting, often used ironically."],
    ["Out of syllabus", "Indian student slang for something unexpected that one was not prepared for."],
  ]),
  brk(),

  H1("PROGRESS REPORT"),
  table([700, 2100, 6230], [
    ["No.", "Date / Period", "Work completed"],
    ["1", "11 August 2026", "Project topic finalised; working folder created; collection of protest posters and their sources began."],
    ["2", "15 August 2026", "Principal literature identified: Hakoköngäs, Halmesvaara & Sakki (2020), “Persuasion Through Bitter Humor”."],
    ["3", "20 August 2026", "Working skeleton prepared: research question, thesis statement and chapter outline drafted."],
    ["4", "Late August 2026", "Key dates and timeline of the protest movement compiled; introduction, literature review and methodology drafted."],
    ["5", "28 August 2026", "Progress Report 1 presented to the guide (presentation covering context, research question, literature, methodology and the first worked example)."],
    ["6", "September 2026", "Corpus finalised at 13 items; close readings of Figures 1–7 completed; emerging patterns identified."],
    ["7", "October 2026", "Data tabulated, graphs prepared; findings, conclusion, suggestions, references and appendices completed; final report compiled and submitted."],
  ]),
  blank(), blank(),
  sigTable2(),
);
function sigTable2() {
  const w = [4515, 4515];
  const nb = { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } };
  const cell = (lines, i) => new TableCell({ width: { size: w[i], type: WidthType.DXA }, borders: nb, children: lines.map((l, k) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 276 }, children: [new TextRun({ text: l, bold: k === 1 })] })) });
  return new Table({ width: { size: 9030, type: WidthType.DXA }, columnWidths: w, rows: [new TableRow({ children: [cell(["____________________", "Student", "Gatha Kalpana Vijay"], 0), cell(["____________________", "Project Guide", "Prof. Dr. Umesh Jagdale"], 1)] })] });
}

// ---------------- DOCUMENT ----------------
const page = { size: { width: 11906, height: 16838 }, margin: { left: 2160, right: 1440, top: 1440, bottom: 1440 } };
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] });

const doc = new Document({
  creator: "Gatha Kalpana Vijay",
  title: "The Unserious Generation on the Most Serious Topic",
  styles: {
    default: { document: { run: { font: FONT, size: SZ } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: SZ, bold: true, color: "000000" }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ...["num", "obj", "pat", "find", "inst"].map(r => ({ reference: r, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] })),
    ],
  },
  sections: [
    { properties: { page: { ...page, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } }, titlePage: true }, footers: { default: footer, first: new Footer({ children: [] }) }, children: prelim },
    { properties: { page: { ...page, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } } }, footers: { default: footer }, children: [...body, ...end] },
  ],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync("../FP_Report.docx", b); console.log("written"); });
