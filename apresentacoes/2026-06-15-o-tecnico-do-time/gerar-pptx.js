const fs = require("fs");
const path = require("path");
const PptxGenJS = require("/Users/rodrigosouza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/pptxgenjs");

const OUT_DIR = __dirname;
const ROOT = path.resolve(__dirname, "../..");
const ASSETS = path.join(ROOT, "assets");
const OUTPUT = path.join(OUT_DIR, "culto-kids-2026-06-15-o-tecnico-do-time.pptx");
const vovoPath = path.join(ASSETS, "personagens", "vovo-docura-slide.png");
const soccerBallPath = path.join(ASSETS, "gerados", "2026-06-08-deus-me-deu-um-proposito", "bola-futebol-nitida.png");
const coachPath = path.join(ASSETS, "gerados", "2026-06-15-o-tecnico-do-time", "vovo-docura-tecnica.png");

const pptx = new PptxGenJS();
pptx.author = "Culto Kids";
pptx.company = "Culto Kids";
pptx.subject = "Uma jogada do Mestre - O tecnico do time";
pptx.title = "Culto Kids - O tecnico do time";
pptx.lang = "pt-BR";
pptx.theme = {
  headFontFace: "Comfortaa",
  bodyFontFace: "Comfortaa",
  lang: "pt-BR"
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";
pptx.margin = 0;

const W = 13.333;
const H = 7.5;
const C = {
  black: "000000",
  white: "FFFFFF",
  blue: "2F6690",
  blue2: "48A6D6",
  lightBlue: "E1F3FC",
  green: "6ED33F",
  lightGreen: "E9F9DF",
  yellow: "FFD23B",
  cream: "FFF8DB",
  pink: "FDECF3",
  coral: "FF6B6B",
  orange: "FF9933",
  purple: "8043E6"
};

function addBg(slide, fill = C.white) {
  slide.background = { color: fill };
  addConfetti(slide);
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 6.86,
    w: W,
    h: 0.64,
    fill: { color: C.lightGreen },
    line: { color: C.lightGreen }
  });
}

function addConfetti(slide) {
  const dots = [
    [0.15, 0.12, 0.58, C.blue],
    [1.05, 0.06, 0.42, C.green],
    [2.05, 0.0, 0.62, C.yellow],
    [3.4, 0.25, 0.34, C.green],
    [9.45, 0.04, 0.72, C.blue],
    [10.85, 0.18, 0.58, C.green],
    [11.95, 0.0, 0.62, C.blue],
    [12.55, 0.35, 0.44, C.green],
    [0.1, 6.92, 0.58, C.blue],
    [1.2, 6.78, 0.38, C.green],
    [3.18, 6.74, 0.62, C.yellow],
    [5.0, 6.7, 0.38, C.green],
    [7.0, 6.84, 0.48, C.yellow],
    [8.62, 6.67, 0.7, C.green],
    [10.52, 6.85, 0.58, C.blue],
    [12.5, 6.72, 0.42, C.green]
  ];
  dots.forEach(([x, y, d, color]) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y,
      w: d,
      h: d,
      fill: { color },
      line: { color }
    });
  });
}

function title(slide, text, opts = {}) {
  slide.addText(text, {
    x: opts.x ?? 0.75,
    y: opts.y ?? 0.74,
    w: opts.w ?? 8.9,
    h: opts.h ?? 0.66,
    fontFace: "Comfortaa",
    fontSize: opts.size ?? 34,
    bold: true,
    color: C.black,
    fit: "shrink",
    margin: 0.03,
    breakLine: false
  });
}

function body(slide, text, x, y, w, h, size = 24, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: "Comfortaa",
    fontSize: size,
    bold: opts.bold ?? false,
    color: C.black,
    fit: "shrink",
    valign: opts.valign ?? "mid",
    breakLine: false,
    margin: 0.08,
    paraSpaceAfterPt: 6,
    bullet: opts.bullet
  });
}

function pill(slide, text, x, y, w, color, size = 23) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.62,
    rectRadius: 0.08,
    fill: { color },
    line: { color }
  });
  slide.addText(text, {
    x: x + 0.08,
    y: y + 0.08,
    w: w - 0.16,
    h: 0.46,
    fontFace: "Comfortaa",
    fontSize: size,
    bold: true,
    color: C.black,
    align: "center",
    fit: "shrink",
    margin: 0
  });
}

function addCard(slide, x, y, w, h, color, heading, text) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color },
    line: { color, width: 1.2 }
  });
  slide.addText(heading, {
    x: x + 0.18,
    y: y + 0.14,
    w: w - 0.36,
    h: 0.38,
    fontFace: "Comfortaa",
    fontSize: 21,
    bold: true,
    color: C.black,
    fit: "shrink",
    margin: 0
  });
  slide.addText(text, {
    x: x + 0.18,
    y: y + 0.62,
    w: w - 0.36,
    h: h - 0.74,
    fontFace: "Comfortaa",
    fontSize: 20,
    color: C.black,
    fit: "shrink",
    margin: 0.03,
    valign: "mid"
  });
}

function addVovo(slide, x, y, w, h) {
  if (fs.existsSync(vovoPath)) {
    slide.addImage({ path: vovoPath, x, y, w, h });
  }
}

function addSoccerBall(slide, x, y, d) {
  if (fs.existsSync(soccerBallPath)) {
    slide.addImage({ path: soccerBallPath, x, y, w: d, h: d });
    return;
  }
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: d,
    h: d,
    fill: { color: C.white },
    line: { color: C.black, width: 1.2 }
  });
}

function addField(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.12,
    fill: { color: "79D94A" },
    line: { color: "3DAA32", width: 1.5 }
  });
  slide.addShape(pptx.ShapeType.line, {
    x: x + w / 2,
    y,
    w: 0,
    h,
    line: { color: C.white, width: 2 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + w / 2 - 0.55,
    y: y + h / 2 - 0.55,
    w: 1.1,
    h: 1.1,
    fill: { color: "79D94A", transparency: 100 },
    line: { color: C.white, width: 2 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.14,
    y: y + h / 2 - 0.75,
    w: 0.7,
    h: 1.5,
    fill: { color: "79D94A", transparency: 100 },
    line: { color: C.white, width: 2 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + w - 0.84,
    y: y + h / 2 - 0.75,
    w: 0.7,
    h: 1.5,
    fill: { color: "79D94A", transparency: 100 },
    line: { color: C.white, width: 2 }
  });
}

function addPlayer(slide, x, y, label, color) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: 0.55,
    h: 0.55,
    fill: { color },
    line: { color: C.black, width: 0.5 }
  });
  slide.addText(label, {
    x: x - 0.48,
    y: y + 0.57,
    w: 1.5,
    h: 0.34,
    fontFace: "Comfortaa",
    fontSize: 20,
    bold: true,
    color: C.black,
    align: "center",
    fit: "shrink",
    margin: 0
  });
}

function addCoach(slide, x, y, opts = {}) {
  const h = opts.h ?? 1.92;
  const w = opts.w ?? h * 0.67;
  if (fs.existsSync(coachPath)) {
    slide.addImage({ path: coachPath, x, y, w, h });
  } else {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.32,
      y,
      w: 0.58,
      h: 0.58,
      fill: { color: C.yellow },
      line: { color: C.black, width: 0.7 }
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.18,
      y: y + 0.58,
      w: 0.86,
      h: 1.05,
      rectRadius: 0.08,
      fill: { color: C.blue },
      line: { color: C.black, width: 0.7 }
    });
    slide.addShape(pptx.ShapeType.line, {
      x: x + 0.44,
      y: y + 0.95,
      w: -0.42,
      h: 0.32,
      line: { color: C.black, width: 2 }
    });
    slide.addShape(pptx.ShapeType.line, {
      x: x + 0.75,
      y: y + 0.95,
      w: 0.5,
      h: -0.28,
      line: { color: C.black, width: 2 }
    });
  }
  slide.addText("técnico", {
    x: x + (w - 1.5) / 2,
    y: y + h + 0.06,
    w: 1.5,
    h: 0.34,
    fontFace: "Comfortaa",
    fontSize: 20,
    bold: true,
    color: C.black,
    align: "center",
    fit: "shrink",
    margin: 0
  });
}

function addClipboard(slide, x, y, w, h, fill = C.white) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: fill },
    line: { color: C.black, width: 1.2 }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + w * 0.34,
    y: y - 0.12,
    w: w * 0.32,
    h: 0.32,
    rectRadius: 0.05,
    fill: { color: C.yellow },
    line: { color: C.black, width: 1 }
  });
  for (let i = 0; i < 4; i++) {
    slide.addShape(pptx.ShapeType.line, {
      x: x + 0.28,
      y: y + 0.55 + i * 0.42,
      w: w - 0.56,
      h: 0,
      line: { color: C.blue, width: 1.3 }
    });
  }
}

function addWhistle(slide, x, y, w) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: w * 0.55,
    rectRadius: 0.12,
    fill: { color: C.yellow },
    line: { color: C.black, width: 1.2 }
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + w * 0.12,
    y: y + w * 0.12,
    w: w * 0.22,
    h: w * 0.22,
    fill: { color: C.white },
    line: { color: C.black, width: 1 }
  });
  slide.addShape(pptx.ShapeType.chevron, {
    x: x + w * 0.75,
    y: y + w * 0.12,
    w: w * 0.42,
    h: w * 0.28,
    rotate: 180,
    fill: { color: C.yellow },
    line: { color: C.black, width: 1.2 }
  });
}

function slideVerse(slide, verse, ref, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.86,
    y: 1.62,
    w: 8.35,
    h: 3.5,
    rectRadius: 0.1,
    fill: { color: fill },
    line: { color: fill }
  });
  body(slide, verse, 1.22, 1.94, 7.66, 2.18, 28, { bold: true });
  pill(slide, ref, 2.4, 4.38, 3.55, C.yellow, 23);
}

function addTarget(slide, x, y, d) {
  [C.coral, C.white, C.blue2, C.white, C.yellow].forEach((color, i) => {
    const dd = d - i * 0.34;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + (d - dd) / 2,
      y: y + (d - dd) / 2,
      w: dd,
      h: dd,
      fill: { color },
      line: { color }
    });
  });
}

function addShapeSpeech(slide, x, y, w, h, fill, lines) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.1,
    fill: { color: fill },
    line: { color: fill }
  });
  let yy = y + 0.34;
  lines.forEach(([line, bold]) => {
    slide.addText(line, {
      x: x + 0.42,
      y: yy,
      w: w - 0.84,
      h: 0.62,
      fontFace: "Comfortaa",
      fontSize: bold ? 31 : 27,
      bold,
      color: C.black,
      align: "center",
      fit: "shrink",
      margin: 0
    });
    yy += bold ? 0.86 : 0.72;
  });
}

function notes(slide, text) {
  slide.addNotes(text);
}

// 1
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  addField(s, 7.0, 1.45, 5.25, 4.15);
  addCoach(s, 9.65, 1.62, { h: 2.72 });
  addPlayer(s, 7.58, 3.75, "time", C.blue2);
  addPlayer(s, 8.85, 2.32, "time", C.green);
  addPlayer(s, 11.18, 4.02, "time", C.orange);
  addSoccerBall(s, 8.9, 3.6, 0.62);
  pill(s, "Aula de segunda - 15/06/2026", 0.74, 0.82, 4.95, C.yellow, 21);
  s.addText("UMA JOGADA DO MESTRE", {
    x: 0.76, y: 1.68, w: 5.75, h: 0.45, fontFace: "Comfortaa", fontSize: 24,
    bold: true, color: C.black, margin: 0, fit: "shrink"
  });
  s.addText("O TÉCNICO\nDO TIME", {
    x: 0.72, y: 2.25, w: 6.25, h: 1.7, fontFace: "Comfortaa", fontSize: 44,
    bold: true, color: C.black, margin: 0.02, fit: "shrink", breakLine: false
  });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.76, y: 4.25, w: 5.95, h: 0.86, rectRadius: 0.08,
    fill: { color: C.lightBlue }, line: { color: C.lightBlue }
  });
  body(s, "Tema central: honrando nossos pastores", 0.96, 4.44, 5.55, 0.42, 23, { bold: true });
  notes(s, "A referência recebida está marcada como 14/06. Esta apresentação foi adaptada para a segunda-feira, 15/06/2026.");
}

// 2
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Oi, amiguinhos!");
  addVovo(s, 0.72, 1.35, 3.25, 4.0);
  addShapeSpeech(s, 4.35, 1.55, 7.55, 3.2, C.pink, [
    ["Hoje vamos aprender a", false],
    ["honrar quem cuida", true],
    ["da nossa vida espiritual.", false]
  ]);
  pill(s, "Pergunta rápida", 4.6, 5.12, 2.45, C.yellow, 21);
  body(s, "Quem ajuda um time a jogar com direção?", 7.28, 5.02, 4.8, 0.82, 24, { bold: true });
  notes(s, "Deixar as crianças responderem: técnico, treinador, capitão, professor. Conectar com o papel pastoral.");
}

// 3
{
  const s = pptx.addSlide();
  addBg(s, C.lightBlue);
  title(s, "Tema do mês", { y: 0.74, w: 4.9 });
  s.addText("UMA JOGADA DO MESTRE", {
    x: 0.85, y: 1.65, w: 6.7, h: 0.75, fontFace: "Comfortaa", fontSize: 32,
    bold: true, color: C.black, fit: "shrink", margin: 0
  });
  addField(s, 7.75, 1.12, 4.45, 4.35);
  addCoach(s, 9.48, 1.42, { h: 2.82 });
  addSoccerBall(s, 8.72, 4.12, 0.68);
  pill(s, "Alvo", 1.0, 3.05, 1.6, C.yellow, 23);
  pill(s, "Time", 2.95, 3.05, 1.6, C.green, 23);
  pill(s, "Direção", 4.9, 3.05, 2.0, C.orange, 23);
  body(s, "Um time precisa de direção.\nA igreja também precisa de cuidado e ensino.", 0.95, 4.05, 6.3, 1.36, 26, { bold: true });
  notes(s, "Retomar o tema mensal e explicar que Deus usa líderes para cuidar da igreja.");
}

// 4
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "Versículo-chave do mês");
  slideVerse(s, "Continuo andando para o alvo,\nonde está o meu prêmio maior,\nque é o chamado de Deus\nem Cristo Jesus.", "Filipenses 3:14", C.lightGreen);
  addTarget(s, 9.75, 1.75, 2.05);
  addSoccerBall(s, 10.32, 4.2, 0.75);
  notes(s, "Ler com a turma e repetir: andando para o alvo.");
}

// 5
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Versículo-base da aula");
  slideVerse(s, "Obedeçam aos seus líderes\ne sejam submissos a eles,\npois zelam pela alma\nde vocês.", "Hebreus 13:17", C.pink);
  addClipboard(s, 9.75, 1.72, 2.25, 3.05, C.cream);
  addWhistle(s, 9.52, 4.98, 1.18);
  notes(s, "Explicar que obedecer aqui é ouvir com respeito a liderança espiritual que cuida da igreja.");
}

// 6
{
  const s = pptx.addSlide();
  addBg(s, C.lightGreen);
  title(s, "Nenhum time vence sem direção");
  addField(s, 0.9, 1.55, 6.35, 4.55);
  addCoach(s, 1.08, 1.62, { h: 2.78 });
  addPlayer(s, 3.35, 2.15, "escuta", C.blue2);
  addPlayer(s, 4.9, 3.95, "aprende", C.orange);
  addPlayer(s, 5.88, 2.95, "joga", C.yellow);
  addSoccerBall(s, 4.15, 3.3, 0.55);
  addCard(s, 7.8, 1.55, 4.35, 1.05, C.white, "O técnico orienta", "mostra o caminho");
  addCard(s, 7.8, 2.82, 4.35, 1.05, C.white, "O técnico corrige", "ajuda a melhorar");
  addCard(s, 7.8, 4.1, 4.35, 1.05, C.white, "O técnico incentiva", "anima o time");
  notes(s, "Usar a imagem do técnico do time para aproximar o conceito pastoral.");
}

// 7
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "O pastor cuida do rebanho");
  addVovo(s, 8.9, 1.55, 2.85, 3.5);
  addCard(s, 0.88, 1.5, 3.25, 1.32, C.lightBlue, "Ensina a Palavra", "ajuda a conhecer Jesus");
  addCard(s, 4.38, 1.5, 3.25, 1.32, C.lightGreen, "Ora pela igreja", "cuida com amor");
  addCard(s, 0.88, 3.18, 3.25, 1.32, C.yellow, "Protege", "avisa sobre perigos");
  addCard(s, 4.38, 3.18, 3.25, 1.32, C.pink, "Prepara o time", "ajuda todos a servir");
  body(s, "Pastores cuidam da nossa vida espiritual.", 1.0, 5.18, 7.05, 0.62, 27, { bold: true });
  notes(s, "Falar de forma simples: pastor não é celebridade; é alguém chamado para cuidar, ensinar e orientar.");
}

// 8
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Honrar não é idolatrar");
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.86, y: 1.54, w: 5.95, h: 1.35, rectRadius: 0.08,
    fill: { color: C.lightGreen }, line: { color: C.lightGreen }
  });
  body(s, "Honrar é ouvir,\nrespeitar e orar.", 1.18, 1.78, 5.3, 0.82, 28, { bold: true });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.86, y: 3.28, w: 5.95, h: 1.42, rectRadius: 0.08,
    fill: { color: C.pink }, line: { color: C.pink }
  });
  body(s, "Não transformamos pastor\nem celebridade.", 1.18, 3.52, 5.3, 0.86, 27, { bold: true });
  addTarget(s, 8.55, 1.62, 2.15);
  addWhistle(s, 8.1, 4.2, 1.25);
  body(s, "Jesus é o centro!", 9.35, 4.48, 2.9, 0.55, 26, { bold: true });
  notes(s, "Este slide protege o equilíbrio bíblico da aula: respeito e gratidão sem idolatria.");
}

// 9
{
  const s = pptx.addSlide();
  addBg(s, C.lightBlue);
  title(s, "Vamos participar?");
  body(s, "Complete em voz alta:", 1.05, 1.55, 5.0, 0.48, 26, { bold: true });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 1.05, y: 2.28, w: 10.95, h: 1.35, rectRadius: 0.08,
    fill: { color: C.white }, line: { color: C.white }
  });
  body(s, "Eu posso honrar meus pastores com...", 1.38, 2.55, 10.25, 0.68, 33, { bold: true });
  pill(s, "RESPEITO", 2.05, 4.12, 2.65, C.yellow, 26);
  pill(s, "ORAÇÃO", 5.22, 4.12, 2.3, C.green, 26);
  pill(s, "OBEDIÊNCIA", 8.05, 4.12, 3.0, C.orange, 25);
  body(s, "Agora diga: Jesus é o nosso Mestre!", 2.35, 5.3, 8.8, 0.6, 27, { bold: true });
  notes(s, "Fazer repetição em grupo e perguntar exemplos práticos de respeito.");
}

// 10
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "Atividade: jogando sem técnico");
  addField(s, 8.05, 1.72, 3.85, 3.45);
  addPlayer(s, 8.52, 2.65, "sem rumo", C.blue2);
  addPlayer(s, 9.65, 3.72, "confuso", C.orange);
  addPlayer(s, 10.72, 2.55, "sozinho", C.yellow);
  addSoccerBall(s, 9.55, 2.75, 0.48);
  addCard(s, 0.9, 1.55, 4.0, 1.16, C.lightBlue, "1. Primeira rodada", "uma criança tenta sem direção");
  addCard(s, 0.9, 2.98, 4.0, 1.16, C.lightGreen, "2. Segunda rodada", "outra recebe instruções");
  addCard(s, 0.9, 4.41, 4.0, 1.16, C.yellow, "3. Conversa final", "direção ajuda o time");
  body(s, "Mensagem da brincadeira:\nDeus usa líderes para nos orientar.", 5.2, 2.25, 2.75, 1.5, 24, { bold: true });
  notes(s, "Montar um mini campo. Comparar bagunça sem técnico com missão cumprida após instruções.");
}

// 11
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Desafio da semana");
  addCard(s, 0.92, 1.5, 3.15, 1.42, C.lightBlue, "Ore", "peça a Deus pelos pastores");
  addCard(s, 4.35, 1.5, 3.15, 1.42, C.lightGreen, "Escute", "preste atenção à Palavra");
  addCard(s, 0.92, 3.34, 3.15, 1.42, C.yellow, "Agradeça", "diga uma palavra de carinho");
  addCard(s, 4.35, 3.34, 3.15, 1.42, C.pink, "Obedeça", "pratique o que aprendeu");
  addClipboard(s, 9.08, 1.72, 2.25, 3.05, C.cream);
  body(s, "Pergunte a Deus:\n“Como posso honrar melhor?”", 1.08, 5.35, 7.3, 0.82, 27, { bold: true });
  notes(s, "Dar um desafio prático e simples para a semana.");
}

// 12
{
  const s = pptx.addSlide();
  addBg(s, C.lightGreen);
  addVovo(s, 0.82, 1.7, 3.2, 3.9);
  title(s, "Hoje eu aprendi:", { x: 4.35, y: 1.05, w: 6.6 });
  body(s, "Deus usa líderes para cuidar da igreja.", 4.45, 2.05, 7.25, 0.65, 29, { bold: true });
  body(s, "Eu posso honrar meus pastores.", 4.45, 3.0, 7.1, 0.6, 30, { bold: true });
  body(s, "Jesus continua sendo o centro!", 4.45, 3.92, 7.0, 0.6, 30, { bold: true });
  pill(s, "Até a próxima!", 6.0, 5.25, 3.25, C.yellow, 27);
  addSoccerBall(s, 10.28, 5.14, 0.72);
  notes(s, "Encerrar com oração curta pelos pastores e pela igreja.");
}

pptx.writeFile({ fileName: OUTPUT });
