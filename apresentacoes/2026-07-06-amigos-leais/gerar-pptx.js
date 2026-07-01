import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { Presentation, PresentationFile } = await import(
  require.resolve("@oai/artifact-tool", {
    paths: [
      path.join(
        os.tmpdir(),
        "codex-presentations",
        process.env.CODEX_THREAD_ID ?? "manual-cultokids",
        "2026-07-06-amigos-leais",
        "tmp",
        "node_modules",
      ),
      "/Users/rodrigosouza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules",
    ],
  })
);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "../..");
const ASSETS = path.join(ROOT, "assets");
const OUT = path.join(
  __dirname,
  "culto-kids-2026-07-06-amigos-leais.pptx",
);
const WORKSPACE = path.join(
  os.tmpdir(),
  "codex-presentations",
  process.env.CODEX_THREAD_ID ?? "manual-cultokids",
  "2026-07-06-amigos-leais",
);
const TMP = path.join(WORKSPACE, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const QA = path.join(TMP, "qa");

const W = 1280;
const H = 720;
const FONT = "Arial";
const SOURCE = "Fonte: @flavia.gregio / Mala de Ideias - Ed. Betel";
const C = {
  bg: "#0078D7",
  darkBlue: "#005FAF",
  navy: "#034C8C",
  white: "#FFFFFF",
  yellow: "#FFD23B",
  green: "#27C468",
  coral: "#FF6B6B",
  pink: "#E95C74",
  lilac: "#B89CFF",
  cream: "#FFF8DB",
  red: "#E53935",
};

const vovoPath = path.join(ASSETS, "personagens", "vovo-docura.png");
const vovoSlidePath = path.join(ASSETS, "personagens", "vovo-docura-slide.png");

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

function addShape(slide, geometry, position, fill, options = {}) {
  return slide.shapes.add({
    geometry,
    position,
    fill,
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    borderRadius: options.borderRadius,
    shadow: options.shadow,
  });
}

function addText(slide, value, position, fontSize, options = {}) {
  const box = addShape(slide, "textbox", position, "none");
  box.text = value;
  box.text.style = {
    typeface: FONT,
    fontSize: Math.max(fontSize, 32),
    bold: options.bold ?? true,
    color: options.color ?? C.white,
    alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "middle",
    wrap: "square",
    insets: options.insets ?? { top: 8, right: 10, bottom: 8, left: 10 },
  };
  return box;
}

function addPanel(slide, position, fill = C.darkBlue, options = {}) {
  return addShape(slide, "roundRect", position, fill, {
    borderRadius: options.borderRadius ?? 24,
    shadow: options.shadow ?? "shadow-md",
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
  });
}

function addConfetti(slide) {
  const dots = [
    [8, 8, 72, C.navy],
    [68, 38, 58, C.green],
    [210, 8, 56, C.yellow],
    [1140, 8, 92, C.navy],
    [1212, 42, 58, C.green],
    [1006, 20, 72, C.yellow],
    [8, 628, 84, C.navy],
    [78, 648, 52, C.green],
    [236, 640, 66, C.yellow],
    [1058, 638, 76, C.navy],
    [1182, 650, 58, C.green],
    [1212, 628, 62, C.yellow],
  ];
  for (const [left, top, size, color] of dots) {
    addShape(slide, "ellipse", { left, top, width: size, height: size }, color);
  }
}

function addBackground(slide) {
  slide.background.fill = C.bg;
  addConfetti(slide);
}

function addTitle(slide, title, subtitle = "") {
  addText(slide, title, { left: 70, top: 50, width: 880, height: 76 }, 48);
  if (subtitle) {
    addText(slide, subtitle, { left: 76, top: 126, width: 910, height: 44 }, 30);
  }
}

function addFooter(slide, index) {
  addText(
    slide,
    `${index}/12`,
    { left: 1118, top: 654, width: 112, height: 44 },
    26,
    { align: "center" },
  );
}

function addNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(Array.isArray(notes) ? notes.join("\n") : notes);
  slide.speakerNotes.setVisible(true);
}

async function addImage(slide, filePath, position, alt, fit = "contain", options = {}) {
  if (!(await exists(filePath))) return false;
  slide.images.add({
    blob: await fs.readFile(filePath),
    contentType: "image/png",
    alt,
    fit,
    position,
    geometry: options.geometry,
    borderRadius: options.borderRadius,
  });
  return true;
}

async function addVovo(slide, position) {
  if (
    await addImage(slide, vovoSlidePath, position, "Vovó Doçura", "contain", {
      geometry: "roundRect",
      borderRadius: 24,
    })
  ) {
    return;
  }
  await addImage(slide, vovoPath, position, "Vovó Doçura", "contain", {
    geometry: "roundRect",
    borderRadius: 24,
  });
}

function addChip(slide, text, left, top, width, fill, options = {}) {
  const height = options.height ?? 58;
  addPanel(slide, { left, top, width, height }, fill, { borderRadius: height / 2 });
  const color = fill === C.yellow || fill === C.cream ? C.navy : C.white;
  addText(
    slide,
    text,
    { left: left + 8, top: top + 6, width: width - 16, height: height - 12 },
    options.fontSize ?? 28,
    {
      align: "center",
      color,
      insets: { top: 0, right: 4, bottom: 0, left: 4 },
    },
  );
}

function addHeart(slide, position, fill = C.coral) {
  addShape(slide, "heart", position, fill, { shadow: "shadow-lg" });
}

function addEnvelope(slide, left, top, width, height, fill = C.cream) {
  addShape(slide, "rect", { left, top, width, height }, fill, {
    borderRadius: 18,
    shadow: "shadow-md",
    line: { style: "solid", fill: C.white, width: 4 },
  });
  addShape(
    slide,
    "chevron",
    { left: left + 16, top: top + 12, width: width - 32, height: height - 24, rotation: 90 },
    C.yellow,
    { line: { style: "solid", fill: C.yellow, width: 1 } },
  );
}

function addShield(slide, left, top, width, height, fill = C.yellow) {
  addShape(slide, "homePlate", { left, top, width, height, rotation: 180 }, fill, {
    shadow: "shadow-md",
    line: { style: "solid", fill: C.white, width: 4 },
  });
}

function addBible(slide, left, top, width, height) {
  addShape(slide, "roundRect", { left: left + 12, top: top + 12, width, height }, C.cream, {
    borderRadius: 18,
    shadow: "shadow-md",
    line: { style: "solid", fill: C.white, width: 3 },
  });
  addShape(slide, "roundRect", { left, top, width, height }, C.navy, {
    borderRadius: 18,
    shadow: "shadow-md",
    line: { style: "solid", fill: C.white, width: 4 },
  });
  addShape(slide, "rect", { left: left + width * 0.42, top: top + 20, width: 12, height: 52 }, C.yellow);
  addShape(slide, "rect", { left: left + width * 0.34, top: top + 40, width: 54, height: 12 }, C.yellow);
  addText(slide, "BÍBLIA", { left: left + 22, top: top + height * 0.58, width: width - 44, height: 42 }, 26, {
    align: "center",
  });
}

function addPrayerHands(slide, left, top) {
  addShape(slide, "arc", { left, top, width: 126, height: 150, rotation: 8 }, C.cream, {
    line: { style: "solid", fill: C.white, width: 5 },
  });
  addShape(slide, "arc", { left: left + 74, top, width: 126, height: 150, rotation: -8 }, C.cream, {
    line: { style: "solid", fill: C.white, width: 5 },
  });
  addShape(slide, "ellipse", { left: left + 80, top: top + 132, width: 38, height: 38 }, C.yellow);
}

function addQuestion(slide, number, question, top, fill) {
  addShape(slide, "ellipse", { left: 116, top, width: 64, height: 64 }, fill);
  addText(slide, String(number), { left: 124, top: top + 8, width: 48, height: 44 }, 30, {
    align: "center",
    color: fill === C.yellow ? C.navy : C.white,
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  addPanel(slide, { left: 212, top: top - 8, width: 880, height: 80 }, number % 2 ? C.navy : C.darkBlue);
  addText(slide, question, { left: 242, top: top + 6, width: 820, height: 48 }, 30);
}

async function cover(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addChip(slide, "AMIGOS SÃO PRESENTES DE DEUS", 72, 54, 560, C.green);
  addText(slide, "06/07/2026", { left: 72, top: 128, width: 310, height: 44 }, 30);
  addText(slide, "AMIGOS\nLEAIS", { left: 70, top: 190, width: 620, height: 158 }, 68);
  addPanel(slide, { left: 72, top: 392, width: 660, height: 88 }, C.navy);
  addText(slide, "Amigo leal ama e cuida", { left: 96, top: 410, width: 612, height: 54 }, 36, {
    align: "center",
  });
  addText(slide, "Base: Davi e Jônatas", { left: 80, top: 512, width: 460, height: 44 }, 30);
  addHeart(slide, { left: 850, top: 98, width: 170, height: 160 });
  addShield(slide, 990, 182, 140, 178, C.yellow);
  addBible(slide, 714, 458, 190, 116);
  await addVovo(slide, { left: 930, top: 342, width: 220, height: 292 });
  addFooter(slide, 1);
  addNotes(slide, "Abra apresentando o tema do mês e perguntando o que uma amizade leal faz.");
}

async function monthlyTheme(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Amigos são presentes de Deus", "Tema mensal de julho");
  addPanel(slide, { left: 96, top: 190, width: 670, height: 166 }, C.navy);
  addText(
    slide,
    "Deus nos ensina a amar, cuidar e permanecer perto dos amigos.",
    { left: 132, top: 218, width: 598, height: 108 },
    38,
    { align: "center" },
  );
  addChip(slide, "AMAR", 126, 436, 170, C.coral);
  addChip(slide, "CUIDAR", 332, 436, 190, C.green);
  addChip(slide, "ORAR", 558, 436, 170, C.yellow);
  await addVovo(slide, { left: 852, top: 194, width: 248, height: 356 });
  addFooter(slide, 2);
  addNotes(slide, "Explique que julho terá quatro histórias bíblicas sobre amizade.");
}

function verse(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "O amigo ama em todo tempo", "Vamos repetir juntos");
  addPanel(slide, { left: 116, top: 184, width: 1048, height: 276 }, C.navy);
  addText(
    slide,
    "“Em todo tempo ama o amigo,\ne na angústia nasce o irmão.”",
    { left: 160, top: 224, width: 960, height: 136 },
    44,
    { align: "center" },
  );
  addText(slide, "Provérbios 17:17", { left: 160, top: 368, width: 960, height: 54 }, 34, {
    align: "center",
  });
  addPanel(slide, { left: 252, top: 512, width: 776, height: 82 }, C.green);
  addText(slide, "Amigo leal ama e permanece perto.", { left: 282, top: 530, width: 716, height: 48 }, 34, {
    align: "center",
  });
  addFooter(slide, 3);
  addNotes(slide, "Repita o versículo em duas partes e explique a palavra angústia como dia difícil.");
}

function davidJonathan(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Davi e Jônatas se tornaram amigos");
  addPanel(slide, { left: 92, top: 178, width: 650, height: 160 }, C.navy);
  addText(
    slide,
    "A Bíblia conta que nasceu uma amizade forte entre eles.",
    { left: 126, top: 206, width: 582, height: 104 },
    38,
    { align: "center" },
  );
  addText(slide, "1 Samuel 18:1-4", { left: 110, top: 382, width: 390, height: 46 }, 30);
  addChip(slide, "amizade", 108, 488, 210, C.green);
  addChip(slide, "alegria", 354, 488, 190, C.yellow);
  addChip(slide, "confiança", 580, 488, 230, C.coral);
  addHeart(slide, { left: 894, top: 174, width: 220, height: 205 }, C.pink);
  addShield(slide, 910, 406, 178, 190, C.yellow);
  addFooter(slide, 4);
  addNotes(slide, "Conte a história com simplicidade: Davi e Jônatas aprenderam a confiar um no outro.");
}

function noJealousy(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Jônatas não teve inveja");
  addText(
    slide,
    "Ele poderia competir com Davi, mas escolheu amar e ajudar.",
    { left: 92, top: 176, width: 760, height: 120 },
    40,
  );
  addPanel(slide, { left: 106, top: 354, width: 816, height: 112 }, C.green);
  addText(
    slide,
    "Amizade leal fica feliz com a bênção do outro.",
    { left: 142, top: 378, width: 744, height: 64 },
    38,
    { align: "center" },
  );
  addShape(slide, "ellipse", { left: 974, top: 188, width: 150, height: 150 }, C.yellow);
  addText(slide, "sem\ninveja", { left: 980, top: 222, width: 138, height: 82 }, 34, {
    align: "center",
    color: C.navy,
  });
  addHeart(slide, { left: 914, top: 420, width: 164, height: 150 }, C.coral);
  addFooter(slide, 5);
  addNotes(slide, "Aplique com exemplos: quando um amigo ganha, aprende algo novo ou recebe elogio.");
}

function loyalProtects(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Amigo leal protege");
  addPanel(slide, { left: 94, top: 174, width: 660, height: 180 }, C.navy);
  addText(
    slide,
    "Jônatas avisou Davi sobre o perigo e fez o que era certo.",
    { left: 130, top: 204, width: 588, height: 116 },
    38,
    { align: "center" },
  );
  addText(slide, "1 Samuel 20:16-17", { left: 112, top: 390, width: 430, height: 46 }, 30);
  addPanel(slide, { left: 160, top: 510, width: 738, height: 78 }, C.darkBlue);
  addText(slide, "Pergunta: como posso proteger um amigo sem brigar?", { left: 188, top: 526, width: 682, height: 48 }, 30, {
    align: "center",
  });
  addShield(slide, 892, 184, 212, 246, C.yellow);
  addText(slide, "cuidado", { left: 872, top: 446, width: 250, height: 58 }, 36, {
    align: "center",
  });
  addFooter(slide, 6);
  addNotes(slide, "Direcione as respostas para atitudes seguras: chamar um adulto, não zombar, defender com respeito.");
}

function loyalDoesGood(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Amigo leal faz o bem");
  const items = [
    ["fala a verdade", 116, 194, C.green],
    ["não zomba", 396, 194, C.coral],
    ["não abandona", 676, 194, C.yellow],
    ["pede perdão", 956, 194, C.pink],
  ];
  for (const [label, left, top, color] of items) {
    addShape(slide, "ellipse", { left, top, width: 142, height: 142 }, color);
    addText(slide, label, { left: left - 52, top: top + 166, width: 246, height: 70 }, 28, {
      align: "center",
      color: C.white,
    });
  }
  addPanel(slide, { left: 166, top: 512, width: 948, height: 86 }, C.navy);
  addText(slide, "Ser leal é escolher o que agrada a Deus.", { left: 198, top: 530, width: 884, height: 50 }, 36, {
    align: "center",
  });
  addFooter(slide, 7);
  addNotes(slide, "Peça que as crianças escolham uma das quatro atitudes para praticar.");
}

async function jesusFriend(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Jesus é o melhor amigo");
  addPanel(slide, { left: 96, top: 182, width: 676, height: 164 }, C.navy);
  addText(
    slide,
    "Jesus nos amou primeiro e nos ensina a amar os amigos.",
    { left: 132, top: 210, width: 604, height: 108 },
    38,
    { align: "center" },
  );
  addText(slide, "João 15:12-13", { left: 112, top: 384, width: 340, height: 46 }, 30);
  addPanel(slide, { left: 146, top: 508, width: 650, height: 78 }, C.green);
  addText(slide, "Jesus ensina a amizade fiel.", { left: 176, top: 524, width: 590, height: 48 }, 34, {
    align: "center",
  });
  await addVovo(slide, { left: 870, top: 192, width: 250, height: 360 });
  addFooter(slide, 8);
  addNotes(slide, "Mantenha Jesus no centro: a amizade de Davi e Jônatas aponta para o amor que Jesus ensina.");
}

function beLoyalToday(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Como ser amigo leal hoje?");
  addChip(slide, "orar", 144, 192, 210, C.green, { height: 82, fontSize: 34 });
  addChip(slide, "respeitar", 402, 192, 250, C.yellow, { height: 82, fontSize: 34 });
  addChip(slide, "ajudar", 700, 192, 220, C.coral, { height: 82, fontSize: 34 });
  addChip(slide, "perdoar", 968, 192, 220, C.pink, { height: 82, fontSize: 34 });
  addPanel(slide, { left: 164, top: 370, width: 952, height: 132 }, C.navy);
  addText(
    slide,
    "Na escola, em casa e na igreja,\nposso cuidar dos meus amigos.",
    { left: 198, top: 396, width: 884, height: 80 },
    38,
    { align: "center" },
  );
  addFooter(slide, 9);
  addNotes(slide, "Peça exemplos concretos para cada palavra: orar, respeitar, ajudar e perdoar.");
}

function activity(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Atividade: amigo secreto de oração");
  addEnvelope(slide, 126, 190, 230, 150);
  addPrayerHands(slide, 910, 186);
  addPanel(slide, { left: 406, top: 180, width: 468, height: 222 }, C.navy);
  addText(
    slide,
    "Cada criança ora por um amigo durante o mês.",
    { left: 438, top: 220, width: 404, height: 118 },
    36,
    { align: "center" },
  );
  addPanel(slide, { left: 224, top: 508, width: 832, height: 82 }, C.green);
  addText(slide, "No final, entregue uma cartinha de encorajamento.", { left: 254, top: 526, width: 772, height: 48 }, 32, {
    align: "center",
  });
  addFooter(slide, 10);
  addNotes(slide, "Organize duplas ou sorteio simples. Garanta que visitantes sejam incluídos.");
}

function review(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Revisão da amizade");
  addQuestion(slide, 1, "Quem eram os amigos da história?", 168, C.green);
  addQuestion(slide, 2, "O que significa ser leal?", 272, C.yellow);
  addQuestion(slide, 3, "Como Jesus nos ensina a amar?", 376, C.coral);
  addQuestion(slide, 4, "Qual atitude vou praticar esta semana?", 480, C.pink);
  addFooter(slide, 11);
  addNotes(slide, "Faça a revisão em voz alta e deixe as crianças responderem com frases curtas.");
}

async function challenge(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Desafio da semana");
  addPanel(slide, { left: 92, top: 174, width: 710, height: 156 }, C.green);
  addText(slide, "Ore por um amigo.", { left: 126, top: 218, width: 642, height: 68 }, 50, {
    align: "center",
  });
  addText(
    slide,
    "Escolha uma pessoa para abençoar com oração e cuidado.",
    { left: 118, top: 382, width: 710, height: 100 },
    36,
    { align: "center" },
  );
  addPanel(slide, { left: 104, top: 548, width: 690, height: 62 }, C.navy);
  addText(slide, "Jesus, ensina-nos a ser amigos leais.", { left: 132, top: 558, width: 634, height: 42 }, 30, {
    align: "center",
  });
  await addVovo(slide, { left: 878, top: 182, width: 250, height: 350 });
  addPanel(slide, { left: 100, top: 626, width: 798, height: 44 }, C.darkBlue);
  addText(slide, SOURCE, { left: 120, top: 631, width: 758, height: 34 }, 24, {
    align: "center",
  });
  addFooter(slide, 12);
  addNotes(slide, "Convide as crianças para uma oração curta e cite a fonte visível no encerramento.");
}

async function writeTextFiles() {
  await fs.mkdir(QA, { recursive: true });
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  await fs.writeFile(
    path.join(TMP, "source-notes.txt"),
    [
      "Culto Kids - 2026-07-06 - Amigos leais",
      "Roteiro local: apresentacoes/2026-07-06-amigos-leais/roteiro.md",
      "Roteiro-base: imagens fornecidas pelo professor a partir da postagem do Instagram @flavia.gregio, com indicação de conteúdo do Livro Mala de Ideias - Ed. Betel.",
      `Fonte curta inserida no slide 12: ${SOURCE}`,
      "Versiculo principal: Proverbios 17:17.",
      "Textos de apoio: 1 Samuel 18:1-4; 1 Samuel 20:16-17, 41-42; Joao 15:12-13.",
      "Assets: assets/personagens/vovo-docura.png e assets/personagens/vovo-docura-slide.png.",
      "Nenhum asset externo novo foi usado.",
      "",
    ].join("\n"),
  );
  await fs.writeFile(
    path.join(TMP, "slide-plan.txt"),
    [
      "Modo: create",
      "Destino: apresentacoes/2026-07-06-amigos-leais/culto-kids-2026-07-06-amigos-leais.pptx",
      "Slide size: 1280 x 720, 16:9.",
      "Paleta: fundo #0078D7; apoio #005FAF/#034C8C; acentos #27C468, #FFD23B, #FF6B6B, #E95C74, #B89CFF.",
      "Tipografia: Arial, minimo 32px em texto visivel; fonte curta com 32px.",
      "Slides: capa, tema mensal, versiculo, Davi e Jonatas, sem inveja, protege, faz o bem, Jesus amigo, aplicacao, atividade, revisao, desafio.",
      "Objetos: textos, shapes, chips, Biblia, coracoes, escudo, envelope, maos em oracao e imagem da Vovo Docura.",
      "",
    ].join("\n"),
  );
}

async function main() {
  await writeTextFiles();

  const presentation = Presentation.create({
    slideSize: { width: W, height: H },
  });

  await cover(presentation);
  await monthlyTheme(presentation);
  verse(presentation);
  davidJonathan(presentation);
  noJealousy(presentation);
  loyalProtects(presentation);
  loyalDoesGood(presentation);
  await jesusFriend(presentation);
  beLoyalToday(presentation);
  activity(presentation);
  review(presentation);
  await challenge(presentation);

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(
      path.join(PREVIEW, `${stem}.png`),
      await presentation.export({ slide, format: "png", scale: 1 }),
    );
    await fs.writeFile(
      path.join(LAYOUT, `${stem}.layout.json`),
      await (await slide.export({ format: "layout" })).text(),
    );
  }

  await writeBlob(
    path.join(PREVIEW, "deck-montage.webp"),
    await presentation.export({ format: "webp", montage: true, scale: 1 }),
  );

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUT);
  await fs.rm(`${OUT}.inspect.ndjson`, { force: true });

  await fs.writeFile(
    path.join(QA, "visual-qa.txt"),
    [
      "# Visual QA",
      "",
      "- PPTX exists and is non-empty: pending command verification.",
      "- Expected slide count: 12.",
      "- Every final slide rendered: yes, see preview/slide-01.png through slide-12.png.",
      "- Contact sheet or montage generated: deck-montage.webp.",
      "- Intended fonts: Arial with minimum visible text >= 32px.",
      "- Source citation visible on slide 12.",
      "- Source notes reviewed: source-notes.txt.",
      "- Known caveat: final human visual review should inspect preview PNGs before Drive publication.",
      "",
    ].join("\n"),
  );

  console.log(`PPTX: ${OUT}`);
  console.log(`Workspace: ${WORKSPACE}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
