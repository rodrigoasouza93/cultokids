import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { Presentation, PresentationFile } = await import(
  require.resolve("@oai/artifact-tool", {
    paths: [
      "/Users/rodrigosouza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules",
    ],
  })
);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "../..");
const ASSETS = path.join(ROOT, "assets");
const OUT = path.join(
  __dirname,
  "culto-kids-2026-06-29-jogando-limpo.pptx",
);
const WORKSPACE = path.join(
  os.tmpdir(),
  "codex-presentations",
  process.env.CODEX_THREAD_ID ?? "manual-cultokids",
  "2026-06-29-jogando-limpo",
);
const TMP = path.join(WORKSPACE, "tmp");
const PREVIEW = path.join(TMP, "preview");
const LAYOUT = path.join(TMP, "layout");
const QA = path.join(TMP, "qa");

const W = 1280;
const H = 720;
const FONT = "Arial";
const C = {
  bg: "#0078D7",
  darkBlue: "#005FAF",
  navy: "#034C8C",
  white: "#FFFFFF",
  yellow: "#FFD23B",
  green: "#27C468",
  red: "#E53935",
  coral: "#FF6B6B",
  pink: "#E95C74",
  lilac: "#B89CFF",
  cream: "#FFF8DB",
};

const vovoPath = path.join(ASSETS, "personagens", "vovo-docura.png");
const vovoSlidePath = path.join(ASSETS, "personagens", "vovo-docura-slide.png");
const ballPath = path.join(
  ASSETS,
  "gerados",
  "2026-06-08-deus-me-deu-um-proposito",
  "bola-futebol-nitida.png",
);

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

function addTitle(slide, title, subtitle = "") {
  addText(slide, title, { left: 70, top: 52, width: 880, height: 72 }, 48);
  if (subtitle) {
    addText(slide, subtitle, { left: 76, top: 126, width: 900, height: 42 }, 28, {
      bold: true,
    });
  }
}

function addConfetti(slide) {
  const dots = [
    [-55, -45, 116, C.navy],
    [70, 42, 62, C.green],
    [205, -36, 72, C.yellow],
    [1125, -42, 116, C.navy],
    [1230, 40, 62, C.green],
    [1018, 18, 78, C.yellow],
    [-48, 650, 104, C.navy],
    [75, 672, 52, C.green],
    [240, 646, 66, C.yellow],
    [1055, 650, 92, C.navy],
    [1180, 658, 58, C.green],
    [1235, 630, 68, C.yellow],
  ];
  for (const [left, top, size, color] of dots) {
    addShape(slide, "ellipse", { left, top, width: size, height: size }, color);
  }
}

function addBackground(slide) {
  slide.background.fill = C.bg;
  addConfetti(slide);
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

async function addBall(slide, position) {
  if (await addImage(slide, ballPath, position, "Bola de futebol", "contain")) {
    return;
  }
  addShape(slide, "ellipse", position, C.white, {
    line: { style: "solid", fill: C.navy, width: 4 },
  });
}

function addCard(slide, label, body, position, fill, iconText = "") {
  addPanel(slide, position, fill);
  if (iconText) {
    addShape(
      slide,
      "ellipse",
      { left: position.left + 28, top: position.top + 28, width: 68, height: 68 },
      C.white,
    );
    addText(
      slide,
      iconText,
      { left: position.left + 30, top: position.top + 34, width: 64, height: 54 },
      34,
      { color: fill, align: "center", insets: { top: 0, right: 0, bottom: 0, left: 0 } },
    );
  }
  addText(
    slide,
    label,
    {
      left: position.left + (iconText ? 112 : 28),
      top: position.top + 26,
      width: position.width - (iconText ? 140 : 56),
      height: 44,
    },
    32,
  );
  addText(
    slide,
    body,
    {
      left: position.left + 28,
      top: position.top + 88,
      width: position.width - 56,
      height: position.height - 108,
    },
    28,
    { bold: true, align: "center" },
  );
}

function addChip(slide, text, left, top, width, fill) {
  addPanel(slide, { left, top, width, height: 58 }, fill, { borderRadius: 28 });
  const color = fill === C.yellow || fill === C.cream ? C.navy : C.white;
  addText(slide, text, { left: left + 8, top: top + 6, width: width - 16, height: 42 }, 27, {
    align: "center",
    color,
    insets: { top: 0, right: 4, bottom: 0, left: 4 },
  });
}

function addNotes(slide, notes) {
  slide.speakerNotes.textFrame.setText(notes);
  slide.speakerNotes.setVisible(true);
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

async function cover(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addChip(slide, "UMA JOGADA DO MESTRE", 72, 54, 360, C.green);
  addText(slide, "29/06/2026", { left: 72, top: 128, width: 310, height: 44 }, 30);
  addText(slide, "JOGANDO\nLIMPO", { left: 70, top: 190, width: 650, height: 160 }, 66);
  addPanel(slide, { left: 72, top: 392, width: 660, height: 88 }, C.navy);
  addText(
    slide,
    "Caráter vale mais que troféu",
    { left: 96, top: 410, width: 612, height: 54 },
    36,
    { align: "center" },
  );
  addText(slide, "Base: 1 João 2:6", { left: 80, top: 512, width: 420, height: 44 }, 28);
  addShape(slide, "rect", { left: 785, top: 116, width: 152, height: 232, rotation: -9 }, C.green, {
    borderRadius: 18,
    shadow: "shadow-lg",
  });
  addShape(slide, "rect", { left: 895, top: 168, width: 152, height: 232, rotation: 9 }, C.red, {
    borderRadius: 18,
    shadow: "shadow-lg",
  });
  await addVovo(slide, { left: 910, top: 318, width: 250, height: 330 });
  await addBall(slide, { left: 742, top: 475, width: 120, height: 120 });
  addFooter(slide, 1);
  addNotes(slide, [
    "Abra perguntando: o que significa jogar limpo?",
    "Explique que a aula fala sobre caráter, não apenas sobre ganhar.",
  ]);
}

function verse(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Versículo-base", "Vamos repetir juntos");
  addPanel(slide, { left: 128, top: 188, width: 1024, height: 258 }, C.navy);
  addText(
    slide,
    "“Quem diz que permanece nele,\nesse deve também andar\nassim como ele andou.”",
    { left: 168, top: 218, width: 944, height: 154 },
    44,
    { align: "center" },
  );
  addText(slide, "1 João 2:6", { left: 168, top: 372, width: 944, height: 48 }, 32, {
    align: "center",
  });
  addPanel(slide, { left: 264, top: 508, width: 752, height: 84 }, C.green);
  addText(
    slide,
    "Andar com Jesus é viver como Jesus ensinou.",
    { left: 292, top: 526, width: 696, height: 48 },
    32,
    { align: "center" },
  );
  addFooter(slide, 2);
  addNotes(slide, [
    "Leia o versículo inteiro.",
    "Repita em três partes: quem permanece, deve andar, como Jesus andou.",
  ]);
}

async function meaning(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "O que é jogar limpo?");
  addCard(
    slide,
    "No jogo",
    "Respeitar as regras e tratar o outro com respeito.",
    { left: 92, top: 184, width: 494, height: 250 },
    C.green,
    "✓",
  );
  addCard(
    slide,
    "Na vida",
    "Fazer o certo mesmo quando ninguém está olhando.",
    { left: 694, top: 184, width: 494, height: 250 },
    C.darkBlue,
    "♥",
  );
  addPanel(slide, { left: 188, top: 506, width: 904, height: 84 }, C.navy);
  addText(
    slide,
    "Pergunta: que atitude mostra que alguém está jogando limpo?",
    { left: 220, top: 522, width: 840, height: 52 },
    31,
    { align: "center" },
  );
  await addBall(slide, { left: 574, top: 358, width: 132, height: 132 });
  addFooter(slide, 3);
  addNotes(slide, "Peça exemplos simples: não empurrar, falar a verdade, pedir desculpas.");
}

async function notEveryWin(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Nem toda vitória agrada a Deus");
  addShape(slide, "rect", { left: 118, top: 182, width: 180, height: 270, rotation: -7 }, C.red, {
    borderRadius: 20,
    shadow: "shadow-lg",
  });
  addText(slide, "FALTA", { left: 122, top: 286, width: 172, height: 58 }, 38, {
    align: "center",
  });
  addText(
    slide,
    "Mentira, briga, inveja, zombaria e trapaça machucam o time.",
    { left: 362, top: 190, width: 640, height: 132 },
    38,
  );
  addPanel(slide, { left: 360, top: 370, width: 672, height: 118 }, C.navy);
  addText(
    slide,
    "Ganhar fazendo o errado não é vitória de verdade.",
    { left: 392, top: 392, width: 608, height: 72 },
    34,
    { align: "center" },
  );
  await addBall(slide, { left: 990, top: 456, width: 128, height: 128 });
  addFooter(slide, 4);
  addNotes(slide, "Conecte com situações do cotidiano: escola, casa, brincadeiras e igreja.");
}

async function heart(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Deus olha para o coração");
  addText(
    slide,
    "Deus não procura estrelas para o seu time.",
    { left: 92, top: 174, width: 720, height: 64 },
    40,
  );
  addPanel(slide, { left: 92, top: 280, width: 760, height: 132 }, C.green);
  addText(
    slide,
    "Deus procura filhos parecidos com Jesus.",
    { left: 124, top: 308, width: 696, height: 76 },
    40,
    { align: "center" },
  );
  addShape(slide, "heart", { left: 918, top: 174, width: 190, height: 178 }, C.coral, {
    shadow: "shadow-lg",
  });
  addText(slide, "caráter", { left: 860, top: 390, width: 300, height: 60 }, 40, {
    align: "center",
  });
  addPanel(slide, { left: 166, top: 502, width: 948, height: 78 }, C.navy);
  addText(
    slide,
    "Caráter vale mais que troféu.",
    { left: 198, top: 518, width: 884, height: 48 },
    38,
    { align: "center" },
  );
  addFooter(slide, 5);
  addNotes(slide, "Explique caráter como aquilo que somos e escolhemos fazer diante de Deus.");
}

async function jesusModel(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Jesus é o nosso modelo");
  addPanel(slide, { left: 94, top: 172, width: 642, height: 166 }, C.navy);
  addText(
    slide,
    "Andar como Jesus é tratar as pessoas com amor, verdade e humildade.",
    { left: 130, top: 196, width: 570, height: 116 },
    36,
    { align: "center" },
  );
  addChip(slide, "AMOR", 110, 416, 180, C.green);
  addChip(slide, "VERDADE", 326, 416, 220, C.yellow);
  addChip(slide, "HUMILDADE", 582, 416, 260, C.coral);
  addPanel(slide, { left: 136, top: 542, width: 648, height: 64 }, C.darkBlue);
  addText(
    slide,
    "Pergunta: como Jesus tratava as pessoas?",
    { left: 156, top: 552, width: 608, height: 44 },
    28,
    { align: "center" },
  );
  await addVovo(slide, { left: 875, top: 188, width: 246, height: 360 });
  addFooter(slide, 6);
  addNotes(slide, "Mostre que Jesus é o padrão do nosso comportamento, não a comparação com os outros.");
}

function greenCard(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Fruto do Espírito", "Cartão verde");
  addShape(slide, "rect", { left: 88, top: 184, width: 168, height: 250, rotation: -6 }, C.green, {
    borderRadius: 22,
    shadow: "shadow-lg",
  });
  addText(slide, "VERDE", { left: 88, top: 282, width: 168, height: 56 }, 34, {
    align: "center",
  });
  const words = [
    ["amor", 338, 190, C.green],
    ["bondade", 590, 190, C.yellow],
    ["domínio próprio", 820, 190, C.coral, 258],
    ["paz", 338, 326, C.navy],
    ["mansidão", 590, 326, C.pink],
    ["fidelidade", 842, 326, C.lilac],
  ];
  for (const [word, left, top, color, width] of words) {
    addChip(slide, word, left, top, width ?? 210, color);
  }
  addPanel(slide, { left: 292, top: 516, width: 780, height: 76 }, C.darkBlue);
  addText(
    slide,
    "Essas atitudes mostram que Deus está trabalhando em nós.",
    { left: 322, top: 528, width: 720, height: 50 },
    30,
    { align: "center" },
  );
  addFooter(slide, 7);
  addNotes(slide, "Leia Gálatas 5:22-23 como apoio se houver tempo.");
}

function redCard(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Falta espiritual", "Cartão vermelho");
  addShape(slide, "rect", { left: 88, top: 184, width: 168, height: 250, rotation: -6 }, C.red, {
    borderRadius: 22,
    shadow: "shadow-lg",
  });
  addText(slide, "PARE", { left: 88, top: 282, width: 168, height: 56 }, 38, {
    align: "center",
  });
  const words = [
    ["mentira", 338, 190, C.red],
    ["briga", 590, 190, C.coral],
    ["inveja", 842, 190, C.pink],
    ["zombaria", 338, 326, C.navy],
    ["orgulho", 590, 326, C.yellow],
    ["trapaça", 842, 326, C.red],
  ];
  for (const [word, left, top, color] of words) addChip(slide, word, left, top, 210, color);
  addPanel(slide, { left: 292, top: 516, width: 780, height: 76 }, C.darkBlue);
  addText(
    slide,
    "Quando erramos, podemos parar, pedir perdão e mudar.",
    { left: 322, top: 528, width: 720, height: 50 },
    30,
    { align: "center" },
  );
  addFooter(slide, 8);
  addNotes(slide, "Evite envergonhar crianças. Aponte para arrependimento e ajuda de Jesus.");
}

async function characterTraining(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Treino do caráter");
  const items = [
    ["falar a verdade", 92, 184, C.green],
    ["pedir desculpas", 392, 184, C.coral],
    ["dividir", 692, 184, C.yellow],
    ["obedecer", 992, 184, C.pink],
  ];
  for (const [label, left, top, color] of items) {
    addShape(slide, "ellipse", { left, top, width: 134, height: 134 }, color);
    addText(slide, label, { left: left - 52, top: top + 160, width: 238, height: 72 }, 28, {
      align: "center",
    });
  }
  addPanel(slide, { left: 182, top: 506, width: 916, height: 86 }, C.navy);
  addText(
    slide,
    "Caráter também precisa ser treinado todos os dias.",
    { left: 214, top: 524, width: 852, height: 50 },
    36,
    { align: "center" },
  );
  await addBall(slide, { left: 566, top: 322, width: 136, height: 136 });
  addFooter(slide, 9);
  addNotes(slide, "Conecte com a aula anterior: treino espiritual aparece em escolhas práticas.");
}

function activity(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Atividade: cartões do juiz");
  addShape(slide, "rect", { left: 136, top: 196, width: 210, height: 300, rotation: -7 }, C.green, {
    borderRadius: 24,
    shadow: "shadow-lg",
  });
  addText(slide, "FRUTO", { left: 154, top: 288, width: 174, height: 94 }, 38, {
    align: "center",
  });
  addShape(slide, "rect", { left: 934, top: 196, width: 210, height: 300, rotation: 7 }, C.red, {
    borderRadius: 24,
    shadow: "shadow-lg",
  });
  addText(slide, "FALTA", { left: 952, top: 288, width: 174, height: 94 }, 38, {
    align: "center",
  });
  addPanel(slide, { left: 410, top: 212, width: 460, height: 224 }, C.navy);
  addText(
    slide,
    "O professor conta uma atitude.\nAs crianças levantam o cartão correto.",
    { left: 438, top: 246, width: 404, height: 132 },
    34,
    { align: "center" },
  );
  addPanel(slide, { left: 342, top: 526, width: 596, height: 70 }, C.darkBlue);
  addText(slide, "Vamos praticar!", { left: 370, top: 538, width: 540, height: 46 }, 34, {
    align: "center",
  });
  addFooter(slide, 10);
  addNotes(slide, [
    "Use os exemplos do roteiro: verdade, zombaria, pedir desculpas, trapaça, ajudar alguém.",
    "As crianças podem levantar mãos se não houver cartões físicos.",
  ]);
}

function review(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Revisão do time");
  const questions = [
    "O que vale mais que troféu?",
    "Quem é o nosso modelo?",
    "Qual cartão mostra fruto do Espírito?",
    "Qual atitude vou treinar esta semana?",
  ];
  questions.forEach((question, index) => {
    const top = 166 + index * 104;
    addShape(slide, "ellipse", { left: 124, top, width: 64, height: 64 }, index % 2 ? C.yellow : C.green);
    addText(slide, String(index + 1), { left: 132, top: top + 8, width: 48, height: 44 }, 30, {
      align: "center",
      color: index % 2 ? C.navy : C.white,
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
    });
    addPanel(slide, { left: 216, top: top - 8, width: 860, height: 80 }, index % 2 ? C.navy : C.darkBlue);
    addText(slide, question, { left: 242, top: top + 6, width: 808, height: 48 }, 30);
  });
  addFooter(slide, 11);
  addNotes(slide, "Faça a revisão em ritmo rápido, deixando as crianças responderem em voz alta.");
}

async function challenge(presentation) {
  const slide = presentation.slides.add();
  addBackground(slide);
  addTitle(slide, "Desafio da semana");
  addPanel(slide, { left: 98, top: 178, width: 710, height: 156 }, C.green);
  addText(
    slide,
    "Nesta semana,\njogue limpo como Jesus.",
    { left: 132, top: 206, width: 642, height: 104 },
    46,
    { align: "center" },
  );
  addText(
    slide,
    "Escolha uma atitude do fruto do Espírito para praticar.",
    { left: 118, top: 394, width: 700, height: 96 },
    36,
    { align: "center" },
  );
  addPanel(slide, { left: 146, top: 552, width: 624, height: 64 }, C.navy);
  addText(
    slide,
    "Oração: Jesus, ajuda-nos a andar como o Senhor andou.",
    { left: 168, top: 556, width: 580, height: 56 },
    32,
    { align: "center" },
  );
  await addVovo(slide, { left: 870, top: 178, width: 260, height: 388 });
  addFooter(slide, 12);
  addNotes(slide, "Convide as crianças para uma oração curta e concreta.");
}

async function writeTextFiles() {
  await fs.mkdir(QA, { recursive: true });
  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(LAYOUT, { recursive: true });
  await fs.writeFile(
    path.join(TMP, "source-notes.txt"),
    [
      "Culto Kids - 2026-06-29 - Jogando limpo",
      "Roteiro local: apresentacoes/2026-06-29-jogando-limpo/roteiro.md",
      "Roteiro-base: referencias/roteiros/2026-06-uma-jogada-do-mestre/roteiro-dia-28-06.jpeg",
      "Versiculo principal: 1 Joao 2:6, fornecido no roteiro-base.",
      "Texto de apoio: Galatas 5:22-23, citado como apoio sobre fruto do Espirito.",
      "Assets: assets/personagens/vovo-docura.png, assets/personagens/vovo-docura-slide.png, assets/gerados/2026-06-08-deus-me-deu-um-proposito/bola-futebol-nitida.png.",
      "Nenhum asset externo novo foi usado.",
      "",
    ].join("\n"),
  );
  await fs.writeFile(
    path.join(TMP, "slide-plan.txt"),
    [
      "Modo: create",
      "Destino: apresentacoes/2026-06-29-jogando-limpo/culto-kids-2026-06-29-jogando-limpo.pptx",
      "Slide size: 1280 x 720, 16:9.",
      "Paleta: fundo #0078D7; apoio #005FAF/#034C8C; acentos #27C468, #E53935, #FFD23B, #FF6B6B, #E95C74.",
      "Tipografia: Arial, minimo 32px em texto visivel, equivalente aproximado a 24 pt; titulos ate 66px.",
      "Slides: capa, versiculo, jogar limpo, vitoria falsa, coracao, Jesus modelo, cartao verde, cartao vermelho, treino, atividade, revisao, desafio.",
      "Objetos: textos, shapes, cartoes, chips, imagem da Vovo Docura e bola.",
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
  verse(presentation);
  await meaning(presentation);
  await notEveryWin(presentation);
  await heart(presentation);
  await jesusModel(presentation);
  greenCard(presentation);
  redCard(presentation);
  await characterTraining(presentation);
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
      "- Contact sheet or montage reviewed: deck-montage.webp generated.",
      "- Intended fonts: Arial set through typeface in text styles.",
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
