const fs = require("fs");
const path = require("path");
const PptxGenJS = require("/Users/rodrigosouza/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/pptxgenjs");

const OUT_DIR = __dirname;
const ROOT = path.resolve(__dirname, "../..");
const ASSETS = path.join(ROOT, "assets");
const OUTPUT = path.join(OUT_DIR, "culto-kids-2026-06-08-deus-me-deu-um-proposito.pptx");
const vovoPath = path.join(ASSETS, "personagens", "vovo-docura-slide.png");
const soccerBallPath = path.join(ASSETS, "gerados", "2026-06-08-deus-me-deu-um-proposito", "bola-futebol-nitida.png");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Culto Kids";
pptx.company = "Culto Kids";
pptx.subject = "Uma jogada do Mestre - Deus me deu um proposito";
pptx.title = "Culto Kids - Deus me deu um proposito";
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
  purple: "8043E6",
  gray: "E6E6E6"
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
    w: opts.w ?? 8.7,
    h: opts.h ?? 0.65,
    fontFace: "Comfortaa",
    fontSize: opts.size ?? 34,
    bold: true,
    color: C.black,
    fit: "shrink",
    margin: 0.03,
    breakLine: false
  });
}

function subtitle(slide, text, x, y, w, h, size = 25) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: "Comfortaa",
    fontSize: size,
    bold: true,
    color: C.black,
    fit: "shrink",
    margin: 0.06
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
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + d * 0.38,
    y: y + d * 0.34,
    w: d * 0.24,
    h: d * 0.24,
    fill: { color: C.black },
    line: { color: C.black }
  });
  [[0.14, 0.18], [0.68, 0.18], [0.13, 0.66], [0.68, 0.66]].forEach(([px, py]) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + d * px,
      y: y + d * py,
      w: d * 0.18,
      h: d * 0.18,
      fill: { color: C.black },
      line: { color: C.black }
    });
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

function notes(slide, text) {
  slide.addNotes(text);
}

// 1
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  addField(s, 7.2, 1.52, 5.2, 4.0);
  addPlayer(s, 7.72, 2.8, "goleiro", C.yellow);
  addPlayer(s, 9.58, 2.1, "atacante", C.blue2);
  addPlayer(s, 10.9, 3.85, "ajuda", C.green);
  addSoccerBall(s, 9.0, 3.42, 0.62);
  pill(s, "Aula de segunda - 08/06/2026", 0.74, 0.82, 4.95, C.yellow, 21);
  s.addText("UMA JOGADA DO MESTRE", {
    x: 0.76, y: 1.68, w: 5.75, h: 0.45, fontFace: "Comfortaa", fontSize: 24,
    bold: true, color: C.black, margin: 0, fit: "shrink"
  });
  s.addText("CADA JOGADOR\nTEM UMA FUNÇÃO", {
    x: 0.72, y: 2.25, w: 6.35, h: 1.7, fontFace: "Comfortaa", fontSize: 39,
    bold: true, color: C.black, margin: 0.02, fit: "shrink", breakLine: false
  });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.76, y: 4.25, w: 5.95, h: 0.82, rectRadius: 0.08,
    fill: { color: C.lightBlue }, line: { color: C.lightBlue }
  });
  body(s, "Tema central: Deus me deu um propósito", 0.96, 4.42, 5.55, 0.42, 24, { bold: true });
  notes(s, "Apresentar o tema do mês e explicar que a aula foi adaptada do roteiro de domingo para a segunda-feira do Culto Kids.");
}

// 2
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Oi, amiguinhos!");
  addVovo(s, 0.72, 1.35, 3.25, 4.0);
  addShapeSpeech(s, 4.35, 1.6, 7.55, 3.1, C.pink, [
    ["Hoje vamos aprender que", false],
    ["Deus deu um propósito", true],
    ["para cada criança!", false]
  ]);
  pill(s, "Pergunta rápida", 4.6, 5.16, 2.45, C.yellow, 21);
  body(s, "Qual posição você gostaria de jogar em um time?", 7.28, 5.06, 4.8, 0.78, 24, { bold: true });
  notes(s, "Deixar algumas crianças responderem: goleiro, atacante, juiz, técnico, torcida.");
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

// 3
{
  const s = pptx.addSlide();
  addBg(s, C.lightBlue);
  title(s, "Tema do mês", { y: 0.74, w: 4.9 });
  s.addText("UMA JOGADA DO MESTRE", {
    x: 0.85, y: 1.65, w: 6.7, h: 0.75, fontFace: "Comfortaa", fontSize: 32,
    bold: true, color: C.black, fit: "shrink", margin: 0
  });
  addField(s, 7.75, 1.15, 4.45, 4.3);
  addSoccerBall(s, 9.48, 2.95, 0.88);
  pill(s, "Alvo", 1.0, 3.05, 1.6, C.yellow, 23);
  pill(s, "Time", 2.95, 3.05, 1.6, C.green, 23);
  pill(s, "Missão", 4.9, 3.05, 2.0, C.orange, 23);
  body(s, "No futebol, o time precisa jogar junto.\nNo Reino de Deus, também servimos juntos.", 0.95, 4.05, 6.25, 1.36, 26, { bold: true });
  notes(s, "Conectar o tema mensal com o interesse das crianças por futebol e Copa.");
}

// 4
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "Versículo-chave do mês");
  slideVerse(s, "Continuo andando para o alvo,\nonde está o meu prêmio maior,\nque é o chamado de Deus\nem Cristo Jesus.", "Filipenses 3:14", C.lightGreen);
  addTarget(s, 9.6, 1.75, 2.05);
  addSoccerBall(s, 10.18, 4.15, 0.75);
  notes(s, "Ler devagar e repetir a frase: andando para o alvo.");
}

function slideVerse(slide, verse, ref, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.86,
    y: 1.72,
    w: 8.18,
    h: 3.25,
    rectRadius: 0.1,
    fill: { color: fill },
    line: { color: fill }
  });
  body(slide, verse, 1.22, 2.02, 7.5, 2.05, 30, { bold: true });
  pill(slide, ref, 2.45, 4.3, 3.25, C.yellow, 24);
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

// 5
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Versículo-base da aula");
  slideVerse(s, "Se todos fossem um só membro,\nonde estaria o corpo?\nAssim, há muitos membros,\nmas um só corpo.", "1 Coríntios 12:19-20", C.pink);
  addTeamDots(s, 9.4, 1.7);
  notes(s, "Explicar: cada parte do corpo tem valor; no time de Deus cada pessoa também tem valor.");
}

function addTeamDots(slide, x, y) {
  const people = [
    [x, y, C.blue2, "1"],
    [x + 1.1, y + 0.55, C.green, "2"],
    [x + 2.0, y, C.yellow, "3"],
    [x + 0.5, y + 1.65, C.orange, "4"],
    [x + 1.58, y + 1.9, C.coral, "5"]
  ];
  people.forEach(([px, py, color, num]) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: px, y: py, w: 0.72, h: 0.72,
      fill: { color }, line: { color: C.black, width: 0.6 }
    });
    slide.addText(num, {
      x: px, y: py + 0.13, w: 0.72, h: 0.3,
      fontFace: "Comfortaa", fontSize: 20, bold: true,
      color: C.black, align: "center", margin: 0
    });
  });
  body(slide, "muitos jogadores\num só time", x - 0.1, y + 3.0, 3.05, 0.8, 24, { bold: true });
}

// 6
{
  const s = pptx.addSlide();
  addBg(s, C.lightGreen);
  title(s, "No futebol, cada um tem função");
  addField(s, 0.9, 1.55, 6.35, 4.55);
  addPlayer(s, 1.32, 3.35, "goleiro", C.yellow);
  addPlayer(s, 3.2, 2.15, "atacante", C.blue2);
  addPlayer(s, 4.75, 3.95, "capitão", C.orange);
  addPlayer(s, 5.85, 2.95, "juiz", C.white);
  addSoccerBall(s, 3.92, 3.25, 0.55);
  addCard(s, 7.8, 1.55, 4.35, 1.05, C.white, "Goleiro", "defende o gol");
  addCard(s, 7.8, 2.82, 4.35, 1.05, C.white, "Atacante", "ataca o time adversário");
  addCard(s, 7.8, 4.1, 4.35, 1.05, C.white, "Juiz", "ajuda o jogo a ser justo");
  notes(s, "Perguntar: o goleiro faz gol o tempo todo? O atacante apita o jogo? Cada um tem uma função.");
}

// 7
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "No Reino, cada um serve");
  addCard(s, 0.88, 1.58, 2.8, 1.62, C.lightBlue, "Alguns cantam", "louvando a Deus");
  addCard(s, 3.95, 1.58, 2.8, 1.62, C.lightGreen, "Outros ajudam", "com amor e alegria");
  addCard(s, 7.02, 1.58, 2.8, 1.62, C.pink, "Outros oram", "pelas pessoas");
  addCard(s, 10.09, 1.58, 2.38, 1.62, C.yellow, "Outros evangelizam", "falando de Jesus");
  body(s, "Ninguém precisa copiar o outro.\nDeus usa cada criança de um jeito especial.", 1.24, 4.1, 7.2, 1.25, 28, { bold: true });
  notes(s, "Reforçar que todo serviço é importante para Deus.");
}

// 8
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Você é importante!");
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.86, y: 1.54, w: 6.35, h: 1.25, rectRadius: 0.08,
    fill: { color: C.pink }, line: { color: C.pink }
  });
  body(s, "Tem criança que pensa:\n“Eu não sou importante.”", 1.18, 1.74, 5.7, 0.82, 27, { bold: true });
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.86, y: 3.15, w: 6.35, h: 1.45, rectRadius: 0.08,
    fill: { color: C.lightGreen }, line: { color: C.lightGreen }
  });
  body(s, "Mas Deus diz:\n“Você tem valor e propósito!”", 1.18, 3.4, 5.75, 0.9, 28, { bold: true });
  addTarget(s, 8.45, 1.46, 2.1);
  addTeamDots(s, 8.1, 3.78);
  notes(s, "Falar com carinho: toda criança é importante para Deus, para a família e para a igreja.");
}

// 9
{
  const s = pptx.addSlide();
  addBg(s, C.lightBlue);
  title(s, "Vamos participar?");
  subtitle(s, "Complete em voz alta:", 1.05, 1.58, 4.9, 0.5, 26);
  s.addShape(pptx.ShapeType.roundRect, {
    x: 1.05, y: 2.35, w: 10.95, h: 1.25, rectRadius: 0.08,
    fill: { color: C.white }, line: { color: C.white }
  });
  body(s, "No time de Deus, eu tenho uma...", 1.38, 2.58, 10.25, 0.68, 34, { bold: true });
  pill(s, "FUNÇÃO!", 4.55, 4.05, 3.0, C.yellow, 32);
  body(s, "Agora diga para o colega:\nDeus me deu um propósito!", 2.25, 5.05, 8.6, 1.05, 27, { bold: true });
  notes(s, "Fazer as crianças repetirem juntas: Deus me deu um propósito.");
}

// 10
{
  const s = pptx.addSlide();
  addBg(s, C.cream);
  title(s, "Atividade: Bolas na cesta");
  addBasket(s, 9.1, 2.15);
  addSoccerBall(s, 8.04, 4.66, 0.55);
  addSoccerBall(s, 10.15, 4.8, 0.55);
  addCard(s, 0.9, 1.55, 3.65, 1.2, C.lightBlue, "1. Dois grupos", "divida a turma em times");
  addCard(s, 0.9, 3.0, 3.65, 1.2, C.lightGreen, "2. Funções", "cada criança participa");
  addCard(s, 0.9, 4.45, 3.65, 1.2, C.yellow, "3. Todos importam", "premie todos no final");
  body(s, "Mensagem da brincadeira:\nO time precisa de todos!", 4.95, 2.05, 3.6, 1.15, 26, { bold: true });
  notes(s, "Atividade do roteiro: Bolas na cesta, dois grupos. Adaptar com bola real ou bolinhas de papel se necessário.");
}

function addBasket(slide, x, y) {
  slide.addShape(pptx.ShapeType.arc, {
    x, y, w: 2.2, h: 1.1,
    line: { color: C.orange, width: 5 }
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.22, y: y + 0.76, w: 1.76, h: 1.45,
    fill: { color: "FFFFFF", transparency: 100 },
    line: { color: C.orange, width: 4 }
  });
  for (let i = 0; i < 4; i++) {
    slide.addShape(pptx.ShapeType.line, {
      x: x + 0.42 + i * 0.42, y: y + 0.82, w: -0.28, h: 1.25,
      line: { color: C.orange, width: 1.5 }
    });
  }
}

// 11
{
  const s = pptx.addSlide();
  addBg(s, C.white);
  title(s, "Desafio da semana");
  addCard(s, 0.92, 1.58, 3.15, 1.42, C.lightBlue, "Em casa", "ajude alguém sem reclamar");
  addCard(s, 4.35, 1.58, 3.15, 1.42, C.lightGreen, "Na igreja", "procure um jeito de servir");
  addCard(s, 0.92, 3.42, 3.15, 1.42, C.yellow, "Na escola", "trate os colegas com amor");
  addCard(s, 4.35, 3.42, 3.15, 1.42, C.pink, "Na oração", "peça a Deus direção");
  body(s, "Pergunte a Deus:\n“Como eu posso ajudar hoje?”", 1.08, 5.45, 6.8, 0.78, 27, { bold: true });
  notes(s, "Dar um desafio prático para a semana.");
}

// 12
{
  const s = pptx.addSlide();
  addBg(s, C.lightGreen);
  addVovo(s, 0.82, 1.7, 3.2, 3.9);
  title(s, "Hoje eu aprendi:", { x: 4.35, y: 1.05, w: 6.6 });
  body(s, "Deus me deu um propósito.", 4.45, 2.15, 6.9, 0.55, 31, { bold: true });
  body(s, "Eu sou importante no time de Deus.", 4.45, 3.05, 7.2, 0.55, 30, { bold: true });
  body(s, "Eu posso servir com alegria!", 4.45, 3.95, 6.8, 0.55, 30, { bold: true });
  pill(s, "Até a próxima!", 6.0, 5.25, 3.25, C.yellow, 27);
  addSoccerBall(s, 10.28, 5.14, 0.72);
  notes(s, "Encerrar com oração curta agradecendo a Deus pelo propósito de cada criança.");
}

pptx.writeFile({ fileName: OUTPUT });
