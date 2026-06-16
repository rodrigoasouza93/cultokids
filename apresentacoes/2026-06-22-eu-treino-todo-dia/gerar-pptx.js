const fs = require("fs/promises");
const path = require("path");

const OUT_DIR = __dirname;
const ROOT = path.resolve(__dirname, "../..");
const ASSETS = path.join(ROOT, "assets");
const OUTPUT = path.join(OUT_DIR, "culto-kids-2026-06-22-eu-treino-todo-dia.pptx");

const ARTIFACT_TOOL_FALLBACK =
  "/Users/rodrigosouza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/.pnpm/@oai+artifact-tool@file+local-deps+-oai-artifact-tool-oai-artifact_tool-2.8.11.tgz/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const vovoCoachPath = path.join(ASSETS, "gerados", "2026-06-15-o-tecnico-do-time", "vovo-docura-tecnica.png");
const soccerBallPath = path.join(ASSETS, "gerados", "2026-06-08-deus-me-deu-um-proposito", "bola-futebol-nitida.png");

const W = 1280;
const H = 720;
const S = 96;
const FONT = "Arial";
const C = {
  bg: "#0078D7",
  darkBlue: "#004FA3",
  navy: "#0369A1",
  white: "#FFFFFF",
  green: "#6ED33F",
  yellow: "#FFD23B",
  coral: "#FF6B6B",
  pink: "#E95C74",
  lilac: "#B89CFF"
};

async function loadArtifactTool() {
  try {
    return await import("@oai/artifact-tool");
  } catch {
    return await import(ARTIFACT_TOOL_FALLBACK);
  }
}

function p(value) {
  return Math.round(value * S);
}

function pos(x, y, w, h) {
  return { left: p(x), top: p(y), width: p(w), height: p(h) };
}

function addBg(slide) {
  slide.background.fill = C.bg;
  addConfetti(slide);
}

function addShape(slide, geometry, x, y, w, h, fill, options = {}) {
  return slide.shapes.add({
    geometry,
    position: pos(x, y, w, h),
    fill,
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    borderRadius: options.borderRadius
  });
}

function addConfetti(slide) {
  const dots = [
    [-0.35, -0.25, 1.05, C.navy],
    [0.70, 0.35, 0.55, C.green],
    [1.85, -0.35, 0.65, C.yellow],
    [11.65, -0.30, 1.05, C.navy],
    [12.55, 0.35, 0.55, C.green],
    [10.65, 0.20, 0.75, C.yellow],
    [-0.25, 6.65, 0.95, C.navy],
    [0.70, 6.85, 0.45, C.green],
    [2.00, 6.52, 0.55, C.green],
    [11.45, 6.55, 0.78, C.navy],
    [12.55, 6.45, 0.55, C.green],
    [10.45, 6.80, 0.55, C.yellow]
  ];
  dots.forEach(([x, y, d, color]) => addShape(slide, "ellipse", x, y, d, d, color));
}

function text(slide, value, x, y, w, h, size, opts = {}) {
  const box = addShape(slide, "textbox", x, y, w, h, "none");
  box.text = value;
  box.text.style = {
    typeface: FONT,
    fontSize: size,
    bold: true,
    color: C.white,
    alignment: opts.align ?? "left",
    verticalAlignment: opts.valign ?? "middle",
    wrap: "square",
    insets: {
      top: opts.margin ?? 4,
      right: opts.margin ?? 4,
      bottom: opts.margin ?? 4,
      left: opts.margin ?? 4
    }
  };
  return box;
}

function panel(slide, x, y, w, h, color = C.darkBlue) {
  return undefined;
}

function pill(slide, value, x, y, w, color, size = 24) {
  panel(slide, x, y, w, 0.58, color);
  text(slide, value, x + 0.12, y + 0.08, w - 0.24, 0.36, size, { align: "center", margin: 0 });
}

function footer(slide, n) {
  return undefined;
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function image(slide, imagePath, x, y, w, h, alt) {
  if (!(await fileExists(imagePath))) return;
  slide.images.add({
    blob: await fs.readFile(imagePath),
    contentType: "image/png",
    alt,
    fit: "contain",
    position: pos(x, y, w, h)
  });
}

async function ball(slide, x, y, d) {
  if (await fileExists(soccerBallPath)) {
    await image(slide, soccerBallPath, x, y, d, d, "Bola de futebol");
    return;
  }
  addShape(slide, "ellipse", x, y, d, d, C.white, { line: { style: "solid", fill: C.darkBlue, width: 1 } });
}

async function vovoCoach(slide, x, y, w, h) {
  await image(slide, vovoCoachPath, x, y, w, h, "Vovó Doçura vestida como técnica");
}

function title(slide, value, opts = {}) {
  text(slide, value, opts.x ?? 0.72, opts.y ?? 0.58, opts.w ?? 11.8, opts.h ?? 0.82, opts.size ?? 44, {
    align: opts.align ?? "left"
  });
}

async function station(presentation, n, titleText, main, body, call, accent) {
  const slide = presentation.slides.add();
  addBg(slide);
  addShape(slide, "ellipse", 0.78, 0.68, 1.10, 1.10, accent);
  text(slide, String(n), 1.08, 0.86, 0.50, 0.45, 34, { align: "center", margin: 0 });
  title(slide, titleText, { x: 2.02, y: 0.72, w: 7.3, size: 44 });
  panel(slide, 0.92, 2.02, 11.45, 1.32);
  text(slide, main, 1.15, 2.25, 8.20, 0.72, 38, { align: "center" });
  text(slide, body, 1.08, 3.56, 8.15, 1.32, 32, { align: "center" });
  panel(slide, 1.05, 5.50, 11.15, 0.88, accent);
  text(slide, call, 1.18, 5.58, 8.35, 0.62, 26, { align: "center" });
  await vovoCoach(slide, 9.95, 2.05, 2.25, 4.25);
}

async function cover(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  pill(slide, "A JOGADA DO MESTRE", 0.70, 0.48, 5.05, C.green, 24);
  text(slide, "Dia 22/06/2026", 0.72, 1.28, 4.6, 0.42, 28);
  text(slide, "EU TREINO\nTODO DIA", 0.70, 1.82, 7.15, 1.62, 54);
  panel(slide, 0.72, 3.86, 7.55, 1.05);
  text(slide, "Quem anda com Jesus se fortalece", 1.00, 4.06, 7.05, 0.60, 30, { align: "center" });
  text(slide, "Base: 1 Timóteo 4:7", 0.80, 5.18, 5.3, 0.52, 26);
  await image(slide, vovoCoachPath, 8.42, 0.78, 4.17, 6.25, "Vovó Doçura vestida como técnica");
  addShape(slide, "ellipse", 10.82, 5.55, 1.30, 1.30, C.yellow);
  await ball(slide, 10.95, 5.68, 1.05);
  footer(slide, 1);
}

async function verse(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Versículo-base", { x: 0.78, y: 0.80, w: 6.0, size: 42 });
  panel(slide, 0.78, 1.70, 8.55, 2.05);
  text(slide, "“Exercita-te pessoalmente\nna piedade.”", 1.15, 1.97, 7.80, 1.18, 44, { align: "center" });
  text(slide, "1 Timóteo 4:7", 1.15, 3.10, 7.80, 0.45, 26, { align: "center" });
  text(slide, "Treinar com Deus também é para todo dia.", 0.95, 4.35, 7.70, 1.05, 36, { align: "center" });
  await vovoCoach(slide, 9.58, 1.42, 2.25, 3.45);
  panel(slide, 9.10, 4.55, 3.20, 1.00, C.green);
  text(slide, "Vamos memorizar juntos!", 9.18, 4.95, 3.05, 0.58, 26, { align: "center" });
  footer(slide, 2);
}

async function compare(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Jogador Não Treina Só No Dia Do Jogo", { x: 0.72, y: 0.58, w: 9.2, size: 39 });
  panel(slide, 0.78, 1.80, 4.35, 2.30, C.navy);
  text(slide, "No esporte:\ntreino diário fortalece o jogador.", 1.08, 2.15, 3.75, 1.35, 30, { align: "center" });
  panel(slide, 5.42, 1.80, 4.35, 2.30);
  text(slide, "Com Jesus:\nrotina com Deus fortalece o coração.", 5.72, 2.10, 3.75, 1.45, 30, { align: "center" });
  panel(slide, 1.05, 4.68, 8.45, 1.12, C.green);
  text(slide, "Cristão também não cresce só indo ao culto.", 1.32, 4.90, 7.90, 0.58, 30, { align: "center" });
  await vovoCoach(slide, 10.05, 1.22, 2.25, 4.45);
  footer(slide, 3);
}

async function hub(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Nosso Treino Espiritual", { x: 1.00, y: 0.72, w: 7.1, size: 44 });
  addShape(slide, "ellipse", 3.95, 2.48, 1.95, 1.95, C.yellow);
  text(slide, "TODO\nDIA", 4.27, 3.00, 1.32, 0.86, 30, { align: "center" });
  const items = [
    ["ORAÇÃO", 0.88, 2.02, 1.62, C.green],
    ["BÍBLIA", 2.58, 4.20, 1.62, C.coral],
    ["OBEDIÊNCIA", 5.50, 4.20, 2.28, C.lilac],
    ["ADORAÇÃO", 7.95, 2.02, 1.82, C.pink],
    ["BONDADE", 4.78, 1.28, 1.62, C.navy]
  ];
  items.forEach(([label, x, y, w, color]) => {
    addShape(slide, "ellipse", x, y, w, 1.62, color);
    text(slide, label, x + 0.06, y + 0.52, w - 0.12, 0.58, 24, { align: "center", margin: 0 });
  });
  panel(slide, 0.82, 5.92, 8.95, 0.62, C.darkBlue);
  text(slide, "Pequenos passos todos os dias nos aproximam de Deus.", 0.88, 6.08, 8.95, 0.62, 26, { align: "center", margin: 0 });
  await vovoCoach(slide, 10.10, 2.02, 2.05, 4.20);
  footer(slide, 4);
}

async function activity(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Atividade: Circuito De Treino Espiritual", { x: 0.72, y: 0.55, w: 11.6, size: 42, align: "center" });
  const steps = [
    ["1", "Oração", C.yellow],
    ["2", "Bíblia", C.green],
    ["3", "Louvor", C.lilac],
    ["4", "Bondade", C.pink]
  ];
  steps.forEach(([num, label, color], i) => {
    const x = 0.88 + i * 3.05;
    addShape(slide, "ellipse", x, 2.05, 1.05, 1.05, color);
    text(slide, num, x + 0.27, 2.24, 0.50, 0.38, 28, { align: "center", margin: 0 });
    panel(slide, x - 0.32, 3.28, 1.70, 0.68);
    text(slide, label, x - 0.30, 3.40, 1.65, 0.48, 26, { align: "center", margin: 0 });
    if (i < steps.length - 1) text(slide, "→", x + 1.55, 2.32, 1.00, 0.60, 42, { align: "center", margin: 0 });
  });
  panel(slide, 1.15, 5.08, 8.90, 1.18, C.navy);
  text(slide, "As crianças passam pelas estações em rotação e praticam uma ação simples em cada uma.", 1.38, 5.20, 8.45, 0.88, 28, { align: "center" });
  await ball(slide, 10.28, 4.65, 1.75);
  footer(slide, 10);
}

async function review(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Revisão Do Time", { x: 0.72, y: 0.58, w: 8.2, size: 46 });
  const qs = [
    "Qual é o versículo de hoje?",
    "Quais são os treinos espirituais?",
    "O que você vai praticar esta semana?"
  ];
  qs.forEach((q, i) => {
    const y = 1.62 + i * 1.25;
    const color = [C.yellow, C.green, C.coral][i];
    addShape(slide, "ellipse", 0.95, y, 0.65, 0.65, color);
    text(slide, String(i + 1), 1.12, y + 0.12, 0.30, 0.28, 24, { align: "center", margin: 0 });
    text(slide, q, 1.82, y - 0.02, 10.20, 0.65, 31);
  });
  panel(slide, 1.25, 5.85, 10.90, 0.80, C.green);
  text(slide, "Quem anda com Jesus se fortalece!", 1.55, 6.02, 10.30, 0.38, 28, { align: "center" });
  await vovoCoach(slide, 9.95, 2.04, 2.25, 4.25);
  footer(slide, 11);
}

async function close(presentation) {
  const slide = presentation.slides.add();
  addBg(slide);
  title(slide, "Desafio Da Semana", { x: 0.75, y: 0.68, w: 6.5, size: 48 });
  panel(slide, 0.85, 1.75, 7.10, 2.20);
  text(slide, "Treine todo dia:\nore, leia ou ouça a Bíblia, obedeça, adore e faça o bem.", 1.25, 2.05, 6.35, 1.35, 31, { align: "center" });
  panel(slide, 0.85, 4.45, 7.10, 1.35, C.navy);
  text(slide, "Jesus, fortalece nosso coração para caminhar contigo todos os dias. Amém.", 1.18, 4.65, 6.45, 0.82, 25, { align: "center" });
  text(slide, "Quem anda com Jesus se fortalece!", 0.95, 6.12, 7.20, 0.44, 28, { align: "center" });
  await image(slide, vovoCoachPath, 8.35, 0.78, 4.25, 6.12, "Vovó Doçura vestida como técnica");
  footer(slide, 12);
}

async function main() {
  const { Presentation, PresentationFile } = await loadArtifactTool();
  const presentation = Presentation.create({
    slideSize: { width: W, height: H }
  });

  await cover(presentation);
  await verse(presentation);
  await compare(presentation);
  await hub(presentation);
  await station(presentation, 5, "Estação Oração", "Falar com Deus", "Conte a Deus o que sente.\nPeça ajuda. Agradeça.", "Pergunta: pelo que podemos orar hoje?", C.yellow);
  await station(presentation, 6, "Estação Bíblia", "Ouvir a Palavra", "A Bíblia mostra quem Deus é\ne como devemos viver.", "Desafio: repetir 1 Timóteo 4:7 juntos.", C.green);
  await station(presentation, 7, "Estação Obediência", "Praticar o que Deus ensina", "Obedecer é treinar o coração\npara escolher o caminho certo.", "Exemplo: obedecer em casa, na igreja e na escola.", C.coral);
  await station(presentation, 8, "Estação Adoração", "Louvar com a vida", "Adorar é cantar, agradecer\ne viver para agradar a Deus.", "Participação: qual louvor fala com seu coração?", C.lilac);
  await station(presentation, 9, "Estação Bondade", "Amar e ajudar", "Quem anda com Jesus aprende\na fazer o bem todos os dias.", "Missão: fazer uma bondade ainda hoje.", C.pink);
  await activity(presentation);
  await review(presentation);
  await close(presentation);

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT);
  console.log(`Apresentação gerada: ${OUTPUT}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
