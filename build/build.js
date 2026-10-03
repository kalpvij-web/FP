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
  C("**IRONY, HUMOR AND GEN Z LANGUAGE AS POLITICAL RESISTANCE IN THE NEET-UG 2026 PROTESTS**", { after: 360 }),
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
  P("This is to certify that **Gatha Kalpana Vijay**, a student of S.Y.B.A. (English Literature), has satisfactorily completed the Field Project entitled **“The Unserious Generation on the Most Serious Topic: Irony, Humor and Gen Z Language as Political Resistance in the NEET-UG 2026 Protests”** under my guidance and supervision during the academic year 2026–27."),
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
  P("I, **Gatha Kalpana Vijay**, a student of S.Y.B.A. (English Literature), Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner, hereby declare that the Field Project Report entitled **“The Unserious Generation on the Most Serious Topic: Irony, Humor and Gen Z Language as Political Resistance in the NEET-UG 2026 Protests”** is my own original work, carried out under the guidance of **Prof. Dr. Umesh Jagdale**, Department of English."),
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
  P("I would like to thank my guide, **Prof. Dr. Umesh Jagdale Sir**, for guiding me throughout this field project. His suggestions during the progress report presentation helped me figure out what exactly my research question was, and gave this project a proper direction."),
  P("I am also thankful to the Principal and the Head of the Department of English, Sangamner Nagarpalika Arts, D.J. Malpani Commerce and B.N. Sarda Science College (Autonomous), Sangamner, for giving me the opportunity to work on this project, and to all the teachers of the English department for their support."),
  P("A special thanks to all the protestors, students, independent journalists and creators whose posters, photographs and posts became the data for this project. This project exists because of their creativity— all of their work has been credited to the original sources."),
  P("Lastly, thank you to my family and friends for their support, and for tolerating me sending them protest memes at 2 a.m. and calling it ‘research’."),
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
  ["Figure 1", "Quoting famous dialogue of Spider-Man from the Marvel Universe", "fig1"],
  ["Figure 2", "Wordplay— the institution’s own name as the punchline", "fig2"],
  ["Figure 3", "Comparing the government to quick-commerce delivery apps", "fig3"],
  ["Figure 4", "Reclaiming Gen Z slang— ‘Baddies are not apolitical’", "fig4"],
  ["Figure 5", "‘Error 404’— borrowing the language of a broken web page", "fig5"],
  ["Figure 6", "Parody of Dora the Explorer’s question to the audience", "fig6"],
  ["Figure 7", "Pun on ‘leak’— parody of sanitary pad advertisements", "fig7"],
  ["Graph 1", "Items collected, by platform", "g1"],
  ["Graph 2", "Rhetorical devices used in the seven analysed items", "g2"],
  ["Graph 3", "What each analysed item was targeting", "g3"],
];
const TABS = [
  ["Table 1.1", "Four dimensions used to analyse each item", "t11"],
  ["Table 3.1", "Key dates and timeline of the protests", "t31"],
  ["Table 4.1", "Items collected, by platform", "t41"],
  ["Table 4.2", "Summary of the seven analysed items", "t42"],
  ["Table 4.3", "How often each rhetorical device was used", "t43"],
  ["Table 4.4", "Where each item borrowed its joke from", "t44"],
  ["Table 5.1", "Findings against the objectives", "t51"],
];
prelim.push(H1("LIST OF FIGURES AND TABLES"), H3("List of Figures and Graphs"));
FIGS.forEach(f => prelim.push(tocLine(`${f[0]}. ${f[1]}`, f[2])));
prelim.push(H3("List of Tables"));
TABS.forEach(t => prelim.push(tocLine(`${t[0]}. ${t[1]}`, t[2])));

// ---------------- MAIN BODY ----------------
const body = [];
const figCap = (t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LINE, after: 0 }, keepNext: true, children: runs(t, { bold: true }) });

// CHAPTER 1
body.push(...CH("CHAPTER 1", "INTRODUCTION AND RESEARCH METHODOLOGY"));
body.push(
  H2("1.1 Introduction"),
  P("When Gen Z of India swarmed the streets of Jantar Mantar, New Delhi, and various parts of India— protesting over NEET-UG paper leaks and failure of the Indian education system— they carried posters quoting Spider-Man, demanding accountability from the government. Their primary focus targeted the recurring failures and irregularities of the Indian education system, including paper leaks and youth unemployment. They demanded the resignation of the Union Education Minister to hold the government accountable. They also demanded compensation and relief for the families of the students who lost their lives by suicide due to exam turmoil. Six student activists from All India Students’ Association (AISA), along with education activist Sonam Wangchuk, went on an indefinite hunger strike."),
  P("Amidst the protest, Gen Z didn’t just march— they memed. They showed up unserious while being serious at the same time. The relatable, trending brainrot memes, quotes, slang, and reels were used to protest and put forth their demands. They used their creativity to connect the protest and brainrot memes they’ve seen all their life. Unlike the ‘traditional’ protests that have happened so far, Gen Z showed up as they are— full of humor, satire, sarcasm, rage, and wisdom. The use of informal language, satire, dark humor, viral reel dances, brainrot memes, and pop culture references fueled the protest. It was an understanding so far that protests can only be done through formal language, political leaders, and mainstream media— Gen Z broke another stereotype, yet again. Even for the Union, the opposition, and everyone around, this was totally out of syllabus. Therefore, this new-age, unique, humorous, and full-of-satire language of Gen Z needs to be studied."),
  P("The research question this project ought to study is ‘What made irony and humor an effective register for political resistance in the NEET-UG 2026 protests, rather than undermining the protest’s seriousness?’ The memes/reels/slangs/language of Gen Z functioned as a rhetorical mode of political resistance, not decoration around the protest. ‘The unserious generation on the most serious topic’."),

  H2("1.2 Statement of the Research Problem"),
  P("Humor and protest are usually not seen together. A poster quoting a cartoon or a superhero film can easily be called childish or ‘unserious’— and that is exactly how Gen Z gets dismissed most of the time. But during the NEET-UG 2026 protests, these were the posters people remembered, shared and reposted the most. So the question is— how does something that looks so unserious end up carrying one of the most serious demands a generation can make? That is the problem this project tries to understand."),

  H2("1.3 Research Question"),
  P("*What made irony and humor an effective register for political resistance in the NEET-UG 2026 protests, rather than undermining the protest’s seriousness?*", { align: AlignmentType.CENTER, noIndent: true }),

  H2("1.4 Thesis Statement"),
  P("The memes, reels, slang and language of Gen Z functioned as a rhetorical mode of political resistance, not decoration around the protest. The humor was the method, not a distraction from it— ‘the unserious generation on the most serious topic’."),

  H2("1.5 Objectives of the Study"),
  num("To collect and document the protest language— posters, memes and slang— used by Gen Z during the NEET-UG 2026 protests.", "obj"),
  num("To identify the rhetorical devices used in them, like irony, satire, parody, intertextuality, wordplay and slang.", "obj"),
  num("To find out which institutional failure each item targets.", "obj"),
  num("To understand who each item is speaking to— the in-group (Gen Z) and the out-group (government, elders, media).", "obj"),
  num("To explain what humor allowed the protestors to say that formal protest language couldn’t, and to help non-Gen Z generations understand this new-age language of Gen Z.", "obj"),

  H2("1.6 Research Design and Approach"),
  P("This is a qualitative and interpretive study. The method used is multimodal discourse analysis— which basically means reading the image and the words of a poster together, as one message, instead of reading them separately. The framework is adapted from Hakoköngäs, Halmesvaara, and Sakki (2020), which is explained in detail in Chapter 2. Their three dimensions— content, form and rhetorical function— have been adapted in this project into four dimensions, shown in Table 1.1."),
  P("Every item analysed in this project has been analysed along these four dimensions: rhetorical device, institutional failure it targets, its implied audience, and its function within the broader protest discourse. Satire, irony, absurdism, parody, and meme-format literacy have been key rhetorical devices when analyzing the various items used as data. Each meme, each slang term, each ‘unserious’ reel has been used to target the institutional failure to demand accountability."),
  tcaption("Table 1.1. Four dimensions used to analyse each item"),
  table([2200, 6830], [
    ["Dimension", "Question asked"],
    ["1. Rhetorical device", "Which device is doing the work— irony, satire, parody, intertextuality, wordplay, absurdism, slang?"],
    ["2. Institutional target", "Which institutional failure is it pointing at?"],
    ["3. Implied audience", "Who is supposed to ‘get it’ (in-group) and who is left outside the joke (out-group)?"],
    ["4. Function", "Why did it work— what did the humor let it say that formal protest language couldn’t?"],
  ]),
  blank(),

  H2("1.7 Sampling Strategy and Sample Size"),
  P("The corpus for this project consists of 13 items collected from various platforms such as Instagram feeds of protestors, independent journalists, and influencers who visited the protest. Another source used to collect the corpus is Pinterest. One photograph has been sourced via BBC News. The photographs and slang are between June 6, 2026, and July 25, 2026. Each item is unique, but together they tie back to the single objective of accountability and justice. Items were selected based on their rhetorical richness— prioritizing content that was highly shared, creatively distinct, especially illustrative of specific devices (satire, absurdism, irony, parody, etc.), or relatable to Gen Z."),
  P("This means the sampling is purposive— items were picked on purpose, for what they could show, and not randomly. Out of the 13 items collected, **7 items have been analysed in detail** in this project (Figures 1–7). Analysing all 13 in the same depth would have made the project far too long, so these seven were chosen because together they cover every major rhetorical device found in the corpus— intertextuality, irony, parody, wordplay, comparative satire, slang reclamation and dark humor. The other 6 items are listed with their sources in Appendix A."),

  H2("1.8 Data Collection Methods"),
  P("The data used in this project is secondary data— all of it was already publicly posted online. The 13 items were collected from:"),
  bullet("Instagram— 3 items, from the feeds of protestors, independent journalists and influencers who visited the protest."),
  bullet("Pinterest— 9 items, from boards where protest photographs were saved and reshared."),
  bullet("BBC News— 1 photograph from their coverage of the protests."),
  P("For every item, the account it was taken from was noted down so that it could be credited properly. Along with this, the article by Hakoköngäs, Halmesvaara, and Sakki (2020) and some books on memes, humor and resistance were read to build the theoretical background."),

  H2("1.9 Scope and Limitations"),
  bullet("Only 7 out of the 13 items have been analysed in detail, so the findings are based on these seven."),
  bullet("The project only looks at posters, memes and slang between June 6 and July 25, 2026. Speeches, chants and reels have not been analysed in depth."),
  bullet("Since the items were chosen on purpose and the number is small, the findings can’t be generalised to every single post from the protest."),
  bullet("The movement is very recent, so there is almost no academic writing on it yet."),
  bullet("Posts on social media get reposted a lot, so the account credited may sometimes be the one who shared it and not the one who made it."),

  H2("1.10 Chapter Scheme"),
  P("Chapter 1 introduces the topic, the research question and the methodology. Chapter 2 covers the theoretical background and the literature review. Chapter 3 gives the profile of the study context— the NEET-UG exam, the NTA and the protest itself, since this project is not attached to any organisation. Chapter 4 is the analysis of the seven items along with tables and graphs. Chapter 5 gives the findings, conclusion and suggestions."),
  brk(),
);

// CHAPTER 2
body.push(...CH("CHAPTER 2", "THEORETICAL BACKGROUND AND REVIEW OF LITERATURE"));
body.push(
  H2("2.1 Theoretical Background"),
  P("Before getting into the literature review, a few ideas need to be explained because they come up again and again in the analysis."),
  H3("Memes"),
  P("Limor Shifman (2014) describes internet memes as digital items that share common content, form and stance, and that are created, copied and changed by many users. In simple words— a meme isn’t one joke, it’s a format everyone already knows and anyone can join in on. Ryan M. Milner (2016) calls this ‘participatory’, and that is exactly what happened in the protests: anyone who knew the format could make a poster."),
  H3("Multimodality"),
  P("Gunther Kress and Theo van Leeuwen (2006) argue that images, colour, layout and fonts carry meaning just like words do. So a protest poster isn’t only its text— the handwriting, the red colour of an ‘error’ message, a printed cartoon character, all of it is part of what it is saying. Multimodal discourse analysis reads all of these together (Kress, 2012)."),
  H3("Intertextuality and Parody"),
  P("Intertextuality, a term given by Julia Kristeva, means that every text is built out of other texts. Linda Hutcheon (1985) describes parody as ‘repetition with critical difference’— something familiar is repeated, but turned around to criticise. When a protestor rewrites Spider-Man’s line or Dora’s question, this is exactly what is happening."),
  H3("Irony"),
  P("Hutcheon (1994) says irony has an ‘edge’— it says one thing and means another, and in doing so it separates the people who get it from the people who don’t. This is why so many of the posters in this project have a dual-layered audience: an in-group that gets the joke instantly, and an out-group that is left outside it."),
  H3("Humor as Resistance"),
  P("Humor being used against power isn’t new. Mikhail Bakhtin (1984) wrote about the ‘carnivalesque’— moments where laughter turns the official world upside down and makes authority look ridiculous. James C. Scott (1990) showed how people with less power use jokes and disguised speech as a kind of weapon, because it is hard to punish someone for ‘just a joke’. Majken Jul Sørensen (2016), who studied humor in political activism, found that it helps activists get attention, build solidarity and weaken the authority of those in power. So humor and resistance aren’t opposites— humor can be one of the strongest forms of resistance."),

  H2("2.2 Review of Literature"),
  P("Multimodal discourse analysis provides the primary methodological foundation for this project. Hakoköngäs, Halmesvaara, and Sakki (2020) applied this framework to analyze 426 internet memes posted by two far-right Finnish groups, arguing that memes function as a distinct mode of political communication— not mere images, but persuasive tools that combine visual and textual elements to construct ideological meaning. Their study examined memes along three dimensions— 1. Content (thematic categories), 2. Form (visual composition, colour, font), and 3. Rhetorical function (the persuasive work each meme performs). They found that irony and humor were deliberately used to crystallize arguments into an easily shareable form, building in-group solidarity while delegitimizing opponents."),
  P("Their study focuses on far-right nationalist rhetoric, this project adapts the same multimodal, three layer method to a very different context— analyzing how Gen Z used irony and humor not to entrench extremist ideology but to sustain a youth-led protest against institutional failure. No existing literature has yet applied this framework to the NEET-UG 2026 Gen Z protests, as the movement is recent enough that academic scholarship has not caught up— this project fills the gap by extending an established method to a new, unstudied case, while helping non-Gen Z generations to understand the new-age language of Gen Z."),

  H2("2.3 Research Gap"),
  P("Most of the existing work on memes and politics is about Western countries, election campaigns, or extremist groups like the ones Hakoköngäs and others studied. There is hardly anything on Indian student movements, and even less on the way Indian Gen Z actually talks— a mix of English, Hindi, Hinglish, internet slang and global pop culture. And on the NEET-UG 2026 protests specifically, there is no academic work at all yet. That is the gap this project tries to fill."),
  brk(),
);

// CHAPTER 3
body.push(...CH("CHAPTER 3", "PROFILE OF THE STUDY CONTEXT"));
body.push(
  P("Since this project is not done with or for any organisation, this chapter gives the profile of the context instead— the exam, the agency responsible for it, and the protest itself.", { noIndent: true }),
  H2("3.1 The NEET-UG Exam"),
  P("NEET-UG (National Eligibility cum Entrance Test– Undergraduate) is the one national entrance exam for getting into undergraduate medical courses in India, like MBBS, BDS and AYUSH courses. More than twenty lakh students write it every year. For most of these students it means years of preparation, coaching fees, and a huge amount of pressure from themselves and their families— which is why a paper leak isn’t just an ‘irregularity’, it messes with lakhs of lives at once."),
  H2("3.2 The National Testing Agency (NTA)"),
  P("The National Testing Agency is the body set up under the Ministry of Education, Government of India, to conduct big entrance exams including NEET-UG. Since the NTA is responsible for setting the paper, conducting the exam and keeping it secure, it became— along with the Union Ministry of Education— the main target of the protests after the leak."),
  H2("3.3 The NEET-UG 2026 Protests"),
  P("After the NEET-UG 2026 paper leak, students all over India came out to demand accountability. Their main demands were— accountability for the leak and for the repeated failures of the education system, the resignation of the Union Education Minister, and compensation and relief for the families of students who died by suicide because of the exam turmoil. The demonstrations at Jantar Mantar were led by the Cockroach Janta Party (CJP) along with various student groups, including the All India Students’ Association (AISA). Even the name of the group leading the protest is satire— which says a lot about the tone of this whole movement."),
  H3("Key Dates and Timeline"),
  tcaption("Table 3.1. Key dates and timeline of the protests"),
  table([2300, 6730], [
    ["Date", "What happened"],
    ["June 6, 2026", "Physical demonstration began at Jantar Mantar, New Delhi, led by Cockroach Janta Party (CJP) and various student groups for demanding accountability and education minister’s resignation."],
    ["June 28, 2026", "Activist Sonam Wangchuk joined the Jantar Mantar protest and began his indefinite hunger strike."],
    ["July 18, 2026", "Police removed a weakened Sonam Wangchuk from the protest site."],
    ["July 20, 2026", "The ‘Sansad Chalo’— March to Parliament took place in New Delhi— resulting in violent clashes with security forces and student brutality."],
    ["July 23, 2026", "Protests expanded into a nationwide agitation across multiple states such as Tamil Nadu and Kerala."],
    ["July 25, 2026", "Union Education Minister ‘Dharmendra Pradhan’ resigned from his education minister position."],
  ]),
  blank(),
  H2("3.4 Gen Z— the Generation Protesting"),
  P("Gen Z is roughly everyone born between the late 1990s and the early 2010s— the first generation that has grown up completely online. Our everyday language comes from Instagram reels, Pinterest boards, memes, Marvel movies, Hindi-dubbed cartoons, Blinkit, and slang like ‘baddie’, ‘brainrot’ and ‘404’. Almost every NEET-UG aspirant belongs to this generation. So when they protested, it only made sense that they protested in the language they know best— and that language spread online just as fast as it did on the streets."),
  brk(),
);

// CHAPTER 4
body.push(...CH("CHAPTER 4", "DATA ANALYSIS AND INTERPRETATION"));
body.push(
  H2("4.1 Introduction"),
  P("Out of the 13 items collected for this project, 7 have been analysed in detail in this chapter (Figures 1–7). Each one is analysed using the four dimensions from Chapter 1— rhetorical device, institutional target, implied audience and function. After the seven analyses, the findings are put into tables and graphs to see what patterns come up."),
  H2("4.2 The Corpus"),
  tcaption("Table 4.1. Items collected, by platform"),
  table([3010, 3010, 3010], [
    ["Platform", "No. of items", "Percentage"],
    ["Pinterest", "9", "69.2%"],
    ["Instagram", "3", "23.1%"],
    ["BBC News", "1", "7.7%"],
    ["**Total**", "**13**", "**100%**"],
  ]),
  chart("g1.png"),
  caption("Graph 1. Items collected, by platform"),
  P("Most of the items— 9 out of 13— came from Pinterest. That itself is interesting, because Pinterest isn’t a news app. It’s where people save outfits, aesthetics and mood boards. Protest posters being saved and reshared right next to all that shows that for Gen Z, the protest wasn’t something separate from everyday online life— it became a part of it.", { before: 120 }),
  H2("4.3 Analysis— Worked Examples"),
  P("Note: 7 out of the 13 items are analysed below.", { noIndent: true, run: { italics: true } }),
);

const items = [
  ["Figure 1. Quoting famous dialogue of Spider-Man from the Marvel Universe", "Source: Instagram @p9xrider and @p9x_puneet_saxena", [
    "A protester is wearing a Spider-Man mask while holding a handwritten poster that says, *‘With great power comes great responsibility. Can’t say the same for the Indian Government.’* This is intertextuality— borrowing a globally recognized line from the pop-culture franchise (Avengers Multiverse and superheroes). Working through irony, since the original Spider-Man line is an earnest moral lesson about responsibility and accountability that comes with having power; here it’s flipped into an accusation. It is also a form of parody, since the government is being cast as a superhero who has the power but has failed to live up to the responsibility the power demands. It targets the government’s failure to act responsibly despite holding power over education policy— especially the mishandling and negligence behind the NEET-UG paper leaks. The poster doesn’t argue policy details; it makes a moral accusation using a maxim everyone already knows. The implied audience for this is dual-layered. First, the in-group being the Gen Zs— raised on Marvel films; thus the reference lands instantly without needing any explanation, which builds cultural fluency and solidarity amongst the protestors. Second, the out-group being the government officials, older generations, mainstream media— the reference may not even register the same way, which itself marks a generational and cultural gap between protestors and the institutions they’re criticizing.",
    "The poster compresses a complex critique— a government has power without accountability— into a single, instantly legible, highly relatable, and shareable line.",
    "Formal protest language (slogans, official statements, petitions) has to stay measured and defensible— it can’t openly mock authority without risking being dismissed as disrespectful or ‘illegitimate’. A meme sidesteps this entirely. By routing the accusation through Spider-Man, the protester makes the same claim— *the government has failed its basic responsibility*— but wrapped in something ‘just a joke’, which makes it harder for censors, easier to share, relatable to almost every Gen Z, and lower-risk for the person posting it. Humor also lowers the barrier to participation; you don’t need policy expertise to make or share a meme, so it turns outrage into something anyone can contribute to instantly. That’s what formal language structurally can’t do— it can state a grievance, but it can’t spread virally, build in-group belonging through shared laughter, and be simultaneously deniable and cutting.",
  ]],
  ["Figure 2. Wordplay— the institution’s own name as the punchline", "Source: Instagram @chief.of.shitposting", [
    "A protester holds a handwritten poster that reads, *‘The 2 things missing in Education System is EDUCATION & SYSTEM.’* The poster is built like a riddle— it sets up a question (what is missing?) and delivers a punchline that is simply the name of the institution itself. This is wordplay working through structural irony; the very words that are supposed to describe the institution become the accusation against it. It targets the total failure of the education system— not one paper leak, but a system that has become hollow at its core, where neither real education nor a functioning system remains. The implied audience is again dual-layered. The in-group, Gen Z, instantly recognizes the ‘two things’ riddle format from memes and reels, and the punchline lands without explanation. The out-group— officials and older generations— cannot easily dismiss it, because the joke is not an exaggeration; it is literally the complaint.",
    "The poster compresses a sweeping critique of an entire institution into a single, self-evident line.",
    "Formal protest language would need statistics, reports, and long arguments to claim that the education system has collapsed. The joke does it in eleven words. By turning the institution’s own name against it, the poster makes the accusation impossible to argue with— while still being ‘just a joke.’",
  ]],
  ["Figure 3. Comparing the government to quick-commerce delivery apps", "Source: Pinterest @yaashnaaaa", [
    "A black placard held up in the crowd reads, *‘Even Blinkit & Flipkart are faster then your answers.’* This is comparative satire combined with intertextuality from consumer culture— Blinkit (10-minute grocery delivery) and Flipkart are apps Gen Z uses every day. The government is being compared, unfavourably, to a delivery app. It targets the delay and silence of the government and the National Testing Agency in answering students after the NEET-UG paper leak— while lakhs of students waited for clarity, no answers came. The implied audience is urban Gen Z, who live on quick-commerce apps and instantly understand the comparison; for the out-group, the comparison exposes how slow and distant the institution looks to the generation it serves.",
    "The poster measures state accountability against the speed Gen Z expects from everyday life, and in that comparison, official delay begins to look absurd.",
    "Formal language could only say ‘the government has not responded in time.’ The meme makes that delay feel ridiculous. A grocery app reaching your door in ten minutes while a national exam body cannot answer in weeks— the absurdity itself becomes the argument.",
  ]],
  ["Figure 4. Reclaiming Gen Z slang— ‘Baddies are not apolitical’", "Source: Pinterest @Lyaaa", [
    "A protester holds up a handwritten poster that says, *‘BADDIES ARE NOT APOLITICAL.’* ‘Baddie’ is Gen Z slang for a confident, stylish young woman— a word usually associated with selfies, fashion, and Instagram aesthetics rather than politics. This is slang reclamation and the subversion of a stereotype. The poster targets the assumption that Gen Z, especially young women, are unserious, apolitical, and only interested in social media— an assumption that institutions rely on to ignore youth anger. The in-group self-identifies with the label proudly; the out-group (elders, media, politicians) is forced to revise its stereotype.",
    "The poster is an identity declaration— it takes the very label used to dismiss Gen Z as ‘unserious’ and turns it into a political stance. In many ways, it is the thesis of this project written on a poster: the unserious generation on the most serious topic.",
    "Formal protest language would have said ‘young women are politically aware.’ The slang version says it in the language of the people it describes, which makes it both a statement and a form of belonging— being a ‘baddie’ and being political are no longer opposites.",
  ]],
  ["Figure 5. ‘Error 404’— borrowing the language of a broken web page", "Source: Pinterest @FAGUN", [
    "A protester holds a poster that reads *‘error 404’* in red, followed by *‘accountability not found.’* This is internet and tech intertextuality— the HTTP 404 error is the page every internet user sees when a link is broken, or the page they are looking for does not exist. It works through deadpan irony; the poster does not shout, it simply ‘reports an error.’ It targets the absence of accountability— no minister, no agency, no official taking responsibility for the NEET-UG paper leak. For the in-group, digitally native Gen Z, ‘404’ is read instantly. For the out-group, the reference may not register at all, which once again marks the generational and cultural gap between protestors and the institutions they’re criticizing.",
    "The poster frames the government as a broken website— students are searching for accountability, and the page simply does not exist.",
    "Formal protest language would say ‘the authorities have failed to take responsibility.’ The 404 joke turns that into something visual, compact, and instantly shareable, using the red ‘error’ colour itself as part of the message.",
  ]],
  ["Figure 6. Parody of Dora the Explorer’s question to the audience", "Source: Pinterest @mydearnikes", [
    "A protester holds a printed poster of Dora the Explorer, with the line, *‘Kya aapko kahi accountability dikh rahi hai?’* (Do you see accountability anywhere?). This is a parody of Dora’s famous question to the audience from the Hindi-dubbed version of the cartoon, mixed with childhood nostalgia and Hinglish. It targets invisible accountability— the government’s evasion after the NEET-UG leak. The implied audience is Gen Z who grew up watching Hindi-dubbed Dora on television; the call-and-response format pulls the viewer into answering the question, just like the cartoon did.",
    "The poster turns the audience into participants. The obvious answer, ‘No!’, is the protest itself— and it treats the government’s evasions like a children’s game, where even a child can see what is missing.",
    "Formal language can state a demand, but it can’t make the audience say the answer out loud. By borrowing a cartoon every Gen Z remembers, the poster turns a political question into a shared memory, and the shared memory into solidarity.",
  ]],
  ["Figure 7. Pun on ‘leak’— parody of sanitary pad advertisements", "Source: Pinterest @Spring potato", [
    "A protester in the crowd holds a handwritten poster that reads, *‘MY PAD GOT BETTER LEAK PROTECTION THAN EDUCATION SYSTEM.’* This is a pun on the word ‘leak’— a parody of sanitary pad advertisements that promise ‘leak protection,’ placed beside the NEET-UG paper leak. It is also dark humor that breaks the taboo around menstruation by bringing it into a public protest. It targets the paper leak and the complete failure of exam security. The in-group, especially young women, instantly gets the joke; the out-group is jolted— both by the comparison and by a taboo subject being made public.",
    "Shock combined with humor makes the poster unforgettable— an everyday product outperforms a national system.",
    "Formal protest language could never make this comparison; it would be considered ‘improper.’ That is exactly why it works— the poster says the education system is so broken that a sanitary pad is more reliable, and it says it in a way no one can unsee.",
  ]],
];
items.forEach(([cap, src, paras]) => {
  body.push(imgPlaceholder(cap.split(".")[0]), figCap(cap), source(src));
  paras.forEach((t, i) => body.push(P(t, { after: i === paras.length - 1 ? 240 : 120 })));
});

body.push(
  H2("4.4 Data in Tables and Graphs"),
  P("After analysing the seven items, the main points of each analysis were put into a table so they could be compared side by side."),
  tcaption("Table 4.2. Summary of the seven analysed items"),
  table([800, 1900, 2000, 2230, 2100], [
    ["Fig.", "Item", "Rhetorical device", "Institutional target", "Implied audience"],
    ["1", "Spider-Man quote", "Intertextuality, irony, parody", "Power without accountability", "Gen Z raised on Marvel vs. officials, elders, media"],
    ["2", "‘Education & System’", "Wordplay, structural irony", "Total failure of the education system", "Gen Z who know the riddle format vs. officials, elders"],
    ["3", "Blinkit & Flipkart", "Comparative satire, intertextuality", "Delay and silence of Govt./NTA", "Urban Gen Z on quick-commerce apps"],
    ["4", "‘Baddies are not apolitical’", "Slang reclamation", "Stereotype of apolitical youth", "Young women vs. elders, media, politicians"],
    ["5", "‘Error 404’", "Tech intertextuality, deadpan irony", "Absence of accountability", "Digitally native Gen Z vs. out-group"],
    ["6", "Dora the Explorer", "Parody, nostalgia, Hinglish", "Invisible accountability", "Gen Z who grew up on Hindi-dubbed Dora"],
    ["7", "Pad ‘leak protection’", "Pun, parody, dark humor", "Paper leak and exam security", "Young women; out-group is jolted"],
  ]),
  blank(),
  tcaption("Table 4.3. How often each rhetorical device was used"),
  table([3500, 2300, 3230], [
    ["Rhetorical device", "No. of items (out of 7)", "Figures"],
    ["Intertextuality", "3", "1, 3, 5"],
    ["Irony", "3", "1, 2, 5"],
    ["Parody", "3", "1, 6, 7"],
    ["Wordplay / pun", "2", "2, 7"],
    ["Comparative satire", "1", "3"],
    ["Slang reclamation", "1", "4"],
    ["Dark humor", "1", "7"],
  ]),
  P("*Most posters use more than one device, so the total is more than 7.*", { noIndent: true, align: AlignmentType.LEFT }),
  chart("g2.png"),
  caption("Graph 2. Rhetorical devices used in the seven analysed items"),
  P("Intertextuality, irony and parody come out on top, used in three items each. What this shows is that most of the humor came from borrowing— Spider-Man, Dora, Blinkit, a 404 page, a pad ad. Gen Z didn’t need to invent new symbols for the protest. They took things everyone already knew and turned them against the government, and because everyone already knew them, the joke landed instantly.", { before: 120 }),
  chart("g3.png"),
  caption("Graph 3. What each analysed item was targeting"),
  P("On the surface, all seven posters are completely different jokes. But when you look at what they’re actually targeting, three of them point at the exact same thing— no accountability— and the rest point at things that are closely connected to it: a system that has collapsed, no answers from the government, the leak itself, and young people being ignored. The jokes are different. The demand is the same.", { before: 120 }),
  tcaption("Table 4.4. Where each item borrowed its joke from"),
  table([1500, 3500, 4030], [
    ["Figure", "Borrowed from", "What was borrowed"],
    ["1", "Marvel / Spider-Man", "A moral line about power and responsibility"],
    ["2", "Meme riddle format", "The set-up and punchline"],
    ["3", "Blinkit and Flipkart", "The speed of a delivery app"],
    ["4", "Gen Z slang", "The label ‘baddie’"],
    ["5", "The internet", "The ‘404— page not found’ error"],
    ["6", "Dora the Explorer (Hindi dub)", "Dora’s question to the audience"],
    ["7", "Sanitary pad ads", "The promise of ‘leak protection’"],
  ]),
  blank(),
  H2("4.5 Emerging Pattern (Figures 1–7)"),
  P("Across the seven items analysed, four patterns appear. First, borrowed formats— Marvel, Dora, delivery apps, 404 pages, and pad advertisements— familiar culture carries the critique. Second, comparison as accusation— Blinkit and a sanitary pad outperform the State, and the absurdity becomes the argument. Third, reclaimed identity— ‘Baddies are not apolitical’ turns the ‘unserious’ label into a political one. Fourth, a single shared target— whatever the joke, every item points back to one demand: accountability."),
  H2("4.6 Humor as Shield, Sword and Community"),
  P("Looking at all seven together, humor seems to be doing three jobs at once. It is a **shield**— because it’s ‘just a joke’, it’s deniable and harder to censor, even though everyone knows it’s completely serious. It is a **sword**— it lets protestors say things formal protest language never could, like comparing the education system to a sanitary pad or a broken website. And it builds **community**— the moment of ‘you get it’ is itself a kind of solidarity. Shared laughter turns individual anger into something collective, and anyone can join in."),
  brk(),
);

// CHAPTER 5
body.push(...CH("CHAPTER 5", "FINDINGS, CONCLUSION AND SUGGESTIONS"));
body.push(
  H2("5.1 Results and Findings"),
  P("Based on the analysis of 7 out of the 13 items collected, these are the main findings:", { noIndent: true }),
  num("None of the posters were ‘just jokes’. Every single one of the seven had a clear institutional target behind the humor.", "find"),
  num("Intertextuality, irony and parody were the most used devices (three items each). Most of the humor was borrowed from things Gen Z already knew— Marvel, Dora, Blinkit, the internet, ads.", "find"),
  num("Whatever the joke, the target was the same— accountability from the government and the NTA.", "find"),
  num("Almost every poster had a dual-layered audience. Gen Z got it instantly, which built solidarity, while officials and elders were left outside the joke— which itself shows the gap between the protestors and the institutions.", "find"),
  num("Humor let the protestors say things formal protest language can’t. It made the accusations short, shareable, low-risk and ‘deniable’, and allowed comparisons that would otherwise be called ‘improper’.", "find"),
  num("Anyone could take part. You don’t need policy knowledge to make or share a meme, so the protest was open to everyone.", "find"),
  num("The protest became part of everyday online life— 9 of the 13 items were found on Pinterest, saved and shared like any other aesthetic post.", "find"),
  num("Gen Z took back its own image. ‘Baddies are not apolitical’ turned the ‘unserious’ stereotype into a political identity.", "find"),
  tcaption("Table 5.1. Findings against the objectives"),
  table([3800, 5230], [
    ["Objective", "What was found"],
    ["1. Document Gen Z protest language", "13 items collected from Instagram, Pinterest and BBC News (June 6– July 25, 2026); 7 analysed in detail."],
    ["2. Identify rhetorical devices", "Intertextuality, irony and parody (3 each), wordplay/pun (2), comparative satire, slang reclamation and dark humor (1 each)."],
    ["3. Identify institutional targets", "All seven point back to accountability-related failures of the government and the NTA."],
    ["4. Understand the audience", "Mostly dual-layered— in-group solidarity, out-group left outside the joke."],
    ["5. Explain what humor made possible", "Deniability, short and shareable messages, easy participation, and ‘improper’ comparisons."],
  ]),
  blank(),
  H2("5.2 Conclusion"),
  P("This project started with one question— *what made irony and humor an effective register for political resistance in the NEET-UG 2026 protests, rather than undermining the protest’s seriousness?*"),
  P("After analysing seven out of the thirteen items collected, the answer is that the humor didn’t make the protest less serious— it is what carried it. Irony let the protestors say one thing and mean another, and left it to the in-group to complete the meaning. Parody and intertextuality borrowed things everyone already knew and turned them against the government. Wordplay and comparisons made the failure of the system look so absurd that it was almost impossible to defend. And slang like ‘baddie’ took the label that was used to dismiss Gen Z and made it political. All of this made the protest language short, shareable, deniable and open to everyone— which is exactly what formal protest language can’t be."),
  P("Each meme served as a tiny revolution. The memes, reels, slang and language of Gen Z were not decoration around the protest— they were a rhetorical mode of political resistance. The movement that started at Jantar Mantar on June 6, 2026 spread across the country, and on July 25, 2026 the Union Education Minister resigned. This project does not claim that memes alone made that happen, but it does show that humor was central to how this movement spoke, spread, and held together."),
  P("Gen Z has its own language of protest— ironic, visual, full of references, and completely serious underneath. The unserious generation showed up on the most serious topic, and it deserves to be taken seriously."),
  H2("5.3 Recommendations and Suggestions"),
  H3("For the government and institutions"),
  bullet("Memes and humorous posters should be treated as real political expression, not as ‘disrespect’. Every joke analysed here has a very specific complaint in it."),
  bullet("Exam security needs to be stronger, and students need clear and quick answers during a crisis— posters like the Blinkit one and the 404 one show that silence itself becomes the target."),
  H3("For media and older generations"),
  bullet("Instead of brushing off Gen Z’s language, try to understand it. Getting the reference is the first step to getting the demand."),
  H3("For educators"),
  bullet("Memes, protest posters and digital language can be studied in English literature and language classes— they are a living form of satire."),
  bullet("Multimodal analysis can be used in classrooms to help students read media more critically."),
  H3("For further research"),
  bullet("The remaining 6 items in the corpus, along with reels, chants, speeches and the Cockroach Janta Party’s online content, can be analysed in future work."),
  bullet("A bigger collection of items could be studied to check if the same patterns hold for the whole movement."),
  bullet("The NEET-UG 2026 protests can be compared with Gen Z protests in other countries."),
  bullet("Interviews with protestors could show what they were actually thinking when they made these posters."),
  brk(),
);

// ---------------- CONCLUDING SECTIONS ----------------
const end = [];
const ref = (t) => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: LINE, after: 120 }, indent: { left: 567, hanging: 567 }, children: runs(t) });
end.push(
  H1("REFERENCES"),
  P("*(APA 7th edition)*", { noIndent: true, align: AlignmentType.LEFT }),
  H3("Books and Articles"),
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
  H3("Primary Sources (Posters and Photographs)"),
  ref("BBC News. (2026). [Photograph from coverage of the NEET-UG 2026 protests]. https://www.bbc.com/news/articles/cvg9w9dxd7lo"),
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
  H2("Appendix A: List of All 13 Items Collected"),
  P("All 13 items collected for this project are listed below with their sources. Items 1–7 are the ones analysed in Chapter 4. Items 8–13 were collected but not analysed in detail in this project.", { noIndent: true }),
  table([700, 1500, 3400, 3430], [
    ["No.", "Platform", "Source", "Status"],
    ["1", "Instagram", "@p9xrider and @p9x_puneet_saxena", "Analysed— Figure 1"],
    ["2", "Instagram", "@chief.of.shitposting", "Analysed— Figure 2"],
    ["3", "Pinterest", "@yaashnaaaa", "Analysed— Figure 3"],
    ["4", "Pinterest", "@Lyaaa", "Analysed— Figure 4"],
    ["5", "Pinterest", "@FAGUN", "Analysed— Figure 5"],
    ["6", "Pinterest", "@mydearnikes", "Analysed— Figure 6"],
    ["7", "Pinterest", "@Spring potato", "Analysed— Figure 7"],
    ["8", "Instagram", "@purva.thakaree", "Collected, not analysed"],
    ["9", "Pinterest", "@Miss.0", "Collected, not analysed"],
    ["10", "Pinterest", "@Dhruv", "Collected, not analysed"],
    ["11", "Pinterest", "@Sania", "Collected, not analysed"],
    ["12", "BBC News", "bbc.com/news/articles/cvg9w9dxd7lo", "Collected, not analysed"],
    ["13", "Pinterest", "@prwmmm.era", "Collected, not analysed"],
  ]),
  blank(),
  H2("Appendix B: Questions Used to Analyse Each Item"),
  num("What is on the poster/meme— the text, the image, the colours?", "inst"),
  num("What rhetorical device is being used— satire, irony, parody, absurdism, intertextuality, wordplay, in-group reference, slang?", "inst"),
  num("Which institutional failure is it targeting?", "inst"),
  num("Who is the implied audience— who gets it (in-group) and who doesn’t (out-group)?", "inst"),
  num("Why did it work— what did the humor let it say that formal protest language couldn’t?", "inst"),
  H2("Appendix C: Gen Z Words Used in This Project"),
  table([2400, 6630], [
    ["Word", "Meaning"],
    ["Baddie", "A confident, stylish young woman. Usually connected to fashion and Instagram aesthetics."],
    ["Brainrot", "Absurd, repetitive, very-online meme content— or the feeling of having watched too much of it."],
    ["Error 404", "The ‘page not found’ message on the internet. Used to say something doesn’t exist."],
    ["Meme", "An image, text or video format that people copy, change and share online."],
    ["Reel", "A short vertical video on Instagram, usually with trending audio."],
    ["Shitposting", "Posting deliberately absurd or low-effort jokes, usually ironically."],
    ["Out of syllabus", "Student slang for something you were not at all prepared for."],
  ]),
  brk(),

  H1("PROGRESS REPORT"),
  table([700, 2100, 6230], [
    ["No.", "Date", "Work done"],
    ["1", "August 11, 2026", "Topic finalised. Started collecting protest posters along with their sources."],
    ["2", "August 15, 2026", "Found the main article for the literature review— Hakoköngäs, Halmesvaara & Sakki (2020), ‘Persuasion Through Bitter Humor’."],
    ["3", "August 20, 2026", "Made the working skeleton— research question, thesis statement and chapter outline."],
    ["4", "Late August 2026", "Put together the key dates and timeline of the protest. Wrote the introduction, literature review and methodology."],
    ["5", "August 28, 2026", "Presented Progress Report 1 to the guide— context, research question, literature review, methodology and the first worked example (Spider-Man poster)."],
    ["6", "September 2026", "Finalised the corpus at 13 items. Decided to analyse 7 of them in detail. Finished the analysis of Figures 1–7 and the emerging pattern."],
    ["7", "October 2026", "Made the tables and graphs, wrote the findings, conclusion and suggestions, added references and appendices, and put the final report together."],
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
