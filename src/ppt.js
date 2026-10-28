// Genera entregables/Propuesta-Mozart-Movitecnica.pptx
// Estilo: fondo carbón, fotos industriales a sangre con degradé, acento naranja, títulos en mayúsculas.
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/9e5afb32-f417-4b4e-8c86-5d1833191dc0_96639213-28e9-4b21-b321-4e5a36445164/pptx/scripts/apply_theme.js");

const OUT = path.join(__dirname, "../entregables/Propuesta-Mozart-Movitecnica.pptx");
const F = (f) => path.join(__dirname, "fotos", f);
const E = (f) => path.join(__dirname, "../entregables", f);

const C = { bg: "141414", card: "1F1F1F", line: "3A3A3A", orange: "F26B1D", white: "FFFFFF", grey: "B5B5B5", dark: "141414", paper: "FFFFFF", ink: "1A1A1A" };
const THEME = {
  name: "Mozart Industrial", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "141414", lt1: "FFFFFF", dk2: "1F1F1F", lt2: "F2F2F2", accent1: "F26B1D", accent2: "B5B5B5", accent3: "3A3A3A", accent4: "FFB27A", accent5: "8A8A8A", accent6: "5C5C5C", hlink: "F26B1D", folHlink: "B5B5B5" },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Propuesta Mozart para Movitécnica";
pres.author = "Mozart";
const W = 13.333, H = 7.5, M = 0.7;

const footer = (dark = true) => [
  { text: { text: "MOZART × MOVITÉCNICA", options: { x: M, y: 7.0, w: 5, h: 0.3, fontSize: 9, charSpacing: 3, color: dark ? "8A8A8A" : "6B6B6B", margin: 0 } } },
];
pres.defineSlideMaster({ title: "OSCURA", background: { color: C.bg }, objects: footer(true), slideNumber: { x: 12.2, y: 6.95, w: 0.45, h: 0.3, fontSize: 9, color: "8A8A8A", align: "right" } });
pres.defineSlideMaster({ title: "CLARA", background: { color: C.paper }, objects: footer(false), slideNumber: { x: 12.2, y: 6.95, w: 0.45, h: 0.3, fontSize: 9, color: "6B6B6B", align: "right" } });

// ---------- helpers ----------
const kicker = (s, t, x = M, y = 0.6, color = C.orange) =>
  s.addText(t.toUpperCase(), { x, y, w: 8, h: 0.3, fontSize: 11, bold: true, charSpacing: 4, color, margin: 0, isTextBox: true });
const title = (s, t, o = {}) =>
  s.addText(t.toUpperCase(), { x: M, y: 0.95, w: 8.5, h: 1.3, fontSize: 36, bold: true, color: C.white, margin: 0, valign: "top", fontFace: "Arial", isTextBox: true, ...o });
const body = (s, t, o = {}) =>
  s.addText(t, { x: M, y: 2.4, w: 6, h: 0.8, fontSize: 15, color: C.grey, margin: 0, valign: "top", isTextBox: true, ...o });
const bg = (s, img, grad = "grad-left.png") => {
  s.addImage({ path: F(img), x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
  if (grad) s.addImage({ path: F(grad), x: 0, y: 0, w: W, h: H });
};
// Esquinas naranjas: motivo visual tomado de la referencia
const corners = (s, x, y, w, h, len = 0.35, color = C.orange) => {
  const L = (x1, y1, x2, y2) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001, line: { color, width: 2.25 } });
  L(x, y, x + len, y); L(x, y, x, y + len);
  L(x + w - len, y, x + w, y); L(x + w, y, x + w, y + len);
  L(x, y + h, x + len, y + h); L(x, y + h - len, x, y + h);
  L(x + w - len, y + h, x + w, y + h); L(x + w, y + h - len, x + w, y + h);
};
const callout = (s, t, x, y, w, h, o = {}) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText(t, { x: x + 0.25, y, w: w - 0.5, h, fontSize: 15, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true, ...o });
};
const card = (s, x, y, w, h, fill = C.card, transparency = 0) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: fill, transparency }, line: { color: C.line, width: 0.75 } });
const numCircle = (s, n, x, y, d = 0.5) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText(String(n), { x, y, w: d, h: d, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
};
const todo = (s, t, x, y, w, h, dark = true) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: dark ? "2A1A10" : "FFF1E8" }, line: { color: C.orange, width: 1.25, dashType: "dash" } });
  s.addText([{ text: "LUIGGI · ", options: { bold: true } }, { text: t }], { x: x + 0.2, y, w: w - 0.4, h, fontSize: 11, color: dark ? "FFB27A" : "B04A0C", valign: "middle", margin: 0, isTextBox: true });
};
const add = (master, section) => pres.addSlide({ masterName: master, sectionTitle: section });

// ---------- 1. Portada ----------
pres.addSection({ title: "Apertura" });
let s = add("OSCURA", "Apertura");
bg(s, "faja-mina.jpg");
kicker(s, "Propuesta de aliado creativo · Octubre 2026", M, 1.2);
s.addText("INGENIERÍA\nQUE SE VE", { x: M, y: 2.0, w: 8, h: 2.6, fontSize: 64, bold: true, color: C.white, margin: 0, valign: "top", isTextBox: true });
body(s, "Web corporativa B2B, video corto para redes y producción audiovisual en campo para Movitécnica.", { y: 4.8, w: 5.6, h: 0.9, color: "DDDDDD" });
corners(s, 0.35, 0.35, W - 0.7, H - 0.7, 0.45);

// ---------- 2. Lo que entendimos ----------
pres.addSection({ title: "Contexto" });
s = add("OSCURA", "Contexto");
s.addImage({ path: F("polipasto.jpg"), x: 7.6, y: 0, w: W - 7.6, h: H, sizing: { type: "cover", w: W - 7.6, h: H } });
kicker(s, "01 · Lo que entendimos");
title(s, "Mantener en movimiento operaciones que no pueden detenerse", { w: 6.4, fontSize: 30 });
const stats = [["46", "años en el mercado peruano"], ["3", "países: Perú, Chile y Colombia"], ["~70%", "de ingresos viene de minería*"], ["5", "líneas de equipos de izaje y transporte"]];
stats.forEach(([n, l], i) => {
  const x = M + (i % 2) * 3.25, y = 2.75 + Math.floor(i / 2) * 1.95;
  card(s, x, y, 3.0, 1.7);
  s.addText(n, { x: x + 0.25, y: y + 0.15, w: 2.6, h: 0.85, fontSize: 40, bold: true, color: C.orange, margin: 0, isTextBox: true });
  s.addText(l, { x: x + 0.25, y: y + 1.0, w: 2.6, h: 0.55, fontSize: 12, color: C.grey, margin: 0, valign: "top", isTextBox: true });
});
s.addText("* Según declaraciones públicas de la empresa.", { x: M, y: 6.6, w: 6, h: 0.25, fontSize: 9, color: "8A8A8A", margin: 0, isTextBox: true });

// ---------- 3. El reto ----------
s = add("OSCURA", "Contexto");
s.addImage({ path: F("tecnico-polea.jpg"), x: 0, y: 0, w: 6.2, h: H, sizing: { type: "cover", w: 6.2, h: H } });
kicker(s, "02 · El reto", 6.9);
title(s, "Nadie compra un puente grúa por impulso", { x: 6.9, w: 5.8, fontSize: 30 });
[["Operaciones", "Quiere ver capacidad técnica y soporte en faena."], ["Compras", "Necesita fichas técnicas y un canal directo para cotizar."], ["Seguridad", "Exige evidencia de protocolos y trayectoria."]].forEach(([h, t], i) => {
  const y = 2.75 + i * 1.25;
  numCircle(s, i + 1, 6.9, y);
  s.addText(h.toUpperCase(), { x: 7.65, y: y - 0.05, w: 5, h: 0.35, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x: 7.65, y: y + 0.32, w: 5, h: 0.5, fontSize: 13, color: C.grey, margin: 0, valign: "top", isTextBox: true });
});

// ---------- 4. Tres frentes ----------
pres.addSection({ title: "Propuesta" });
s = add("OSCURA", "Propuesta");
kicker(s, "03 · Nuestra propuesta");
title(s, "Tres frentes, una sola línea de marca", { w: 11 });
[["muestra-home-movitecnica.png", E, "Web corporativa B2B", "Catálogo con botón Cotizar en cada equipo."],
 ["muestra-storyboard-reel.png", E, "Video corto para redes", "Reels verticales con material real de faena."],
 ["camarografo.jpg", F, "Audiovisual en campo", "Video institucional, casos y fotografía técnica."]].forEach(([img, fn, h, t], i) => {
  const x = M + i * 4.1, y = 2.3, w = 3.75;
  card(s, x, y, w, 4.25);
  s.addImage({ path: fn(img), x: x + 0.15, y: y + 0.15, w: w - 0.3, h: 2.2, sizing: { type: "cover", w: w - 0.3, h: 2.2 } });
  s.addText(`0${i + 1}`, { x: x + 0.3, y: y + 2.55, w: 1, h: 0.3, fontSize: 11, bold: true, charSpacing: 3, color: C.orange, margin: 0, isTextBox: true });
  s.addText(h.toUpperCase(), { x: x + 0.3, y: y + 2.9, w: w - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x: x + 0.3, y: y + 3.35, w: w - 0.6, h: 0.6, fontSize: 12, color: C.grey, margin: 0, valign: "top", isTextBox: true });
});

// ---------- 5. Quiénes somos (clara) ----------
s = add("CLARA", "Propuesta");
s.addImage({ path: F("cable-acero.jpg"), x: 8.3, y: 0, w: W - 8.3, h: H, sizing: { type: "cover", w: W - 8.3, h: H } });
kicker(s, "04 · Quiénes somos");
title(s, "Oficio de red internacional, estructura ágil", { w: 7, fontSize: 30, color: C.ink });
[["Experiencia", "Equipo formado en TBWA y Publicis."], ["Reconocimientos", "Trabajos premiados en Cannes Lions y New York Festivals."], ["Trato directo", "Quien presenta la propuesta, la ejecuta."]].forEach(([h, t], i) => {
  const y = 2.75 + i * 1.15;
  numCircle(s, i + 1, M, y);
  s.addText(h.toUpperCase(), { x: M + 0.75, y: y - 0.05, w: 6, h: 0.35, fontSize: 15, bold: true, color: C.ink, margin: 0, isTextBox: true });
  s.addText(t, { x: M + 0.75, y: y + 0.32, w: 6.3, h: 0.45, fontSize: 13, color: "4A4A4A", margin: 0, valign: "top", isTextBox: true });
});
todo(s, "Confirmar premio, año y pieza exactos según las Credenciales 2026.", M, 6.25, 7.0, 0.5, false);

// ---------- 6. Casos ----------
s = add("OSCURA", "Propuesta");
kicker(s, "05 · Experiencia relevante");
title(s, "Marcas que confiaron en nosotros", { w: 11 });
[["CLBS", "Web corporativa y rebranding"], ["ELECTRODATA", "Rebranding: experto, profesional, cercano"]].forEach(([h, t], i) => {
  const x = M + i * 6.1, w = 5.8;
  s.addShape(pres.shapes.RECTANGLE, { x, y: 2.3, w, h: 2.9, fill: { color: C.card }, line: { color: C.line, width: 0.75 } });
  corners(s, x + 0.15, 2.45, w - 0.3, 2.6, 0.3);
  s.addText("Imagen del caso", { x, y: 3.5, w, h: 0.4, fontSize: 12, color: "6B6B6B", align: "center", margin: 0, isTextBox: true });
  s.addText(h, { x, y: 5.4, w, h: 0.4, fontSize: 18, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x, y: 5.8, w, h: 0.35, fontSize: 13, color: C.grey, margin: 0, isTextBox: true });
});
s.addText([{ text: "También: ", options: { color: C.grey } }, { text: "BTG PACTUAL · SERRA · LYNCH", options: { bold: true, color: C.white } }], { x: M, y: 6.35, w: 6, h: 0.35, fontSize: 12, margin: 0, isTextBox: true });
todo(s, "Insertar imágenes de CLBS y Electrodata.", 7.6, 6.3, 5.0, 0.45);

// ---------- 7. Concepto ----------
pres.addSection({ title: "Concepto" });
s = add("OSCURA", "Concepto");
bg(s, "puente-grua.jpg");
kicker(s, "06 · Concepto", M, 1.3);
s.addText("INGENIERÍA\nQUE SE VE", { x: M, y: 1.8, w: 7, h: 2.4, fontSize: 60, bold: true, color: C.white, margin: 0, valign: "top", isTextBox: true });
callout(s, "Su trabajo ocurre en faena, de noche y en paradas de planta. Proponemos mostrarlo.", M, 4.45, 5.6, 0.95);
["Mostrar el proceso", "Hablar en datos", "La seguridad como valor"].forEach((t, i) => {
  card(s, M + i * 2.05, 5.7, 1.9, 0.55, "000000", 35);
  s.addText(t, { x: M + i * 2.05, y: 5.7, w: 1.9, h: 0.55, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
});

// ---------- 8. Web B2B ----------
pres.addSection({ title: "Frentes" });
s = add("OSCURA", "Frentes");
kicker(s, "Frente 01 · Web corporativa B2B");
title(s, "Una web que genera cotizaciones", { w: 4.6, fontSize: 30 });
["Cotizar en cada equipo", "Ficha técnica descargable", "Versión PE · CL · CO", "Posventa en primer nivel"].forEach((t, i) => {
  const y = 2.65 + i * 0.85;
  numCircle(s, i + 1, M, y, 0.45);
  s.addText(t, { x: M + 0.65, y, w: 3.8, h: 0.45, fontSize: 15, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true });
});
s.addImage({ path: E("muestra-home-movitecnica.png"), x: 5.55, y: 1.0, w: 7.3, h: 4.56 });
corners(s, 5.4, 0.85, 7.6, 4.86, 0.35);
s.addText("Mockup referencial de la home", { x: 5.55, y: 5.85, w: 7.3, h: 0.3, fontSize: 10, color: "8A8A8A", margin: 0, isTextBox: true });

// ---------- 9. Video redes ----------
s = add("OSCURA", "Frentes");
kicker(s, "Frente 02 · Video corto para redes");
title(s, "Reels que muestran el trabajo real", { w: 11.5, fontSize: 30, h: 0.6 });
s.addImage({ path: E("muestra-storyboard-reel.png"), x: M, y: 1.8, w: 8.5, h: 4.78 });
callout(s, "30 s · 9:16\nSubtítulos siempre activos", 9.55, 1.8, 3.1, 1.1, { fontSize: 14 });
["En faena", "Explicado por un técnico", "Seguridad", "Las personas"].forEach((t, i) => {
  card(s, 9.55, 3.15 + i * 0.85, 3.1, 0.65);
  s.addText(t, { x: 9.8, y: 3.15 + i * 0.85, w: 2.7, h: 0.65, fontSize: 13, bold: true, color: C.white, valign: "middle", margin: 0, isTextBox: true });
});

// ---------- 10. Audiovisual en campo ----------
s = add("OSCURA", "Frentes");
bg(s, "camarografo.jpg");
kicker(s, "Frente 03 · Audiovisual en campo", M, 1.0);
s.addText("GRABAMOS DONDE\nOCURRE EL TRABAJO", { x: M, y: 1.4, w: 7, h: 1.7, fontSize: 40, bold: true, color: C.white, margin: 0, valign: "top", isTextBox: true });
[["Video institucional", "2 a 3 minutos"], ["Casos de cliente", "con aprobación del cliente final"], ["Fotografía técnica", "equipos, procesos y personas"], ["Tomas con dron", "con permiso de la unidad"]].forEach(([h, t], i) => {
  const x = M + (i % 2) * 3.1, y = 3.55 + Math.floor(i / 2) * 1.3;
  card(s, x, y, 2.9, 1.1, "000000", 30);
  s.addText(h.toUpperCase(), { x: x + 0.2, y: y + 0.15, w: 2.6, h: 0.4, fontSize: 13, bold: true, color: C.orange, margin: 0, isTextBox: true });
  s.addText(t, { x: x + 0.2, y: y + 0.55, w: 2.6, h: 0.4, fontSize: 12, color: "DDDDDD", margin: 0, isTextBox: true });
});

// ---------- 11. Seguridad ----------
s = add("OSCURA", "Frentes");
s.addImage({ path: F("charla-seguridad.jpg"), x: 7.3, y: 0, w: W - 7.3, h: H, sizing: { type: "cover", w: W - 7.3, h: H } });
kicker(s, "Protocolo de seguridad en campo");
title(s, "Primero la seguridad, después la cámara", { w: 6.2, fontSize: 30 });
[["Antes", "Inducción y plan de rodaje aprobado por la unidad."], ["Durante", "EPP completo y acompañamiento de Movitécnica."], ["Después", "Revisión del material antes de publicar."]].forEach(([h, t], i) => {
  const y = 2.7 + i * 1.05;
  numCircle(s, i + 1, M, y);
  s.addText(h.toUpperCase(), { x: M + 0.75, y: y - 0.05, w: 5.5, h: 0.35, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x: M + 0.75, y: y + 0.32, w: 5.6, h: 0.45, fontSize: 13, color: C.grey, margin: 0, valign: "top", isTextBox: true });
});
todo(s, "Contratar SCTR para cada persona del equipo antes de entrar a faena.", M, 5.85, 6.1, 0.42);
todo(s, "Conseguir piloto de dron con registro vigente y seguro.", M, 6.37, 6.1, 0.42);

// ---------- 12. Cómo trabajamos ----------
pres.addSection({ title: "Cierre" });
s = add("OSCURA", "Cierre");
kicker(s, "Cómo trabajamos");
title(s, "Un proceso claro en cuatro fases", { w: 11 });
s.addShape(pres.shapes.LINE, { x: M + 0.3, y: 3.15, w: 11.3, h: 0, line: { color: C.line, width: 1.5 } });
[["Inmersión", "Marca, catálogo y prioridades por país."], ["Estrategia y diseño", "Arquitectura web, línea visual y plan de contenidos."], ["Producción", "Desarrollo web y jornadas de rodaje."], ["Lanzamiento", "Publicación, medición y ajustes mensuales."]].forEach(([h, t], i) => {
  const x = M + i * 3.0;
  numCircle(s, i + 1, x + 0.05, 2.9);
  s.addText(h.toUpperCase(), { x, y: 3.7, w: 2.7, h: 0.4, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x, y: 4.15, w: 2.6, h: 0.8, fontSize: 13, color: C.grey, margin: 0, valign: "top", isTextBox: true });
});
todo(s, "Definir plazos por fase (en semanas).", M, 5.8, 5.0, 0.45);

// ---------- 13. Inversión ----------
s = add("OSCURA", "Cierre");
kicker(s, "Inversión");
title(s, "Paquetes por frente, por separado o en conjunto", { w: 11, fontSize: 30 });
[["Frente 01 · Proyecto", "Web corporativa B2B", "S/ ____________", ["Diseño UX/UI", "Cotizar por equipo", "PE · CL · CO", "SEO técnico"]],
 ["Frente 02 · Mensual", "Video corto para redes", "S/ ________ /mes", ["Plan mensual", "Reels con subtítulos", "Carruseles técnicos", "Reporte mensual"]],
 ["Frente 03 · Por jornada", "Audiovisual en campo", "S/ ______ /jornada", ["Rodaje con protocolo", "Fotografía técnica", "Cortes para redes", "Dron opcional"]]].forEach(([k, h, p, items], i) => {
  const x = M + i * 4.1, y = 2.3, w = 3.75;
  card(s, x, y, w, 4.1);
  s.addText(k.toUpperCase(), { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 0.3, fontSize: 10, bold: true, charSpacing: 2, color: C.orange, margin: 0, isTextBox: true });
  s.addText(h.toUpperCase(), { x: x + 0.3, y: y + 0.65, w: w - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(p, { x: x + 0.3, y: y + 1.2, w: w - 0.6, h: 0.6, fontSize: 22, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(items.map((t, j) => ({ text: t, options: { bullet: { code: "25A0" }, breakLine: j < items.length - 1 } })), { x: x + 0.3, y: y + 2.05, w: w - 0.6, h: 1.8, fontSize: 13, color: C.grey, paraSpaceAfter: 6, margin: 0, valign: "top", isTextBox: true });
});
s.addText("Montos en soles. IGV y viáticos a faena: ____________", { x: M, y: 6.55, w: 8, h: 0.3, fontSize: 11, color: "8A8A8A", margin: 0, isTextBox: true });

// ---------- 14. Cierre ----------
s = add("OSCURA", "Cierre");
bg(s, "tecnico-polea.jpg", "grad-full.png");
kicker(s, "Próximos pasos", M, 1.2);
s.addText("CONVERSEMOS", { x: M, y: 1.65, w: 8, h: 1.2, fontSize: 60, bold: true, color: C.white, margin: 0, isTextBox: true });
[["Reunión de 30 minutos", "Prioridades y alcance por país."], ["Propuesta ajustada", "Alcance y montos finales."], ["Arranque", "Inmersión y primera jornada en campo."]].forEach(([h, t], i) => {
  const y = 3.3 + i * 0.95;
  numCircle(s, i + 1, M, y);
  s.addText(h.toUpperCase(), { x: M + 0.75, y: y - 0.05, w: 5, h: 0.35, fontSize: 15, bold: true, color: C.white, margin: 0, isTextBox: true });
  s.addText(t, { x: M + 0.75, y: y + 0.3, w: 5, h: 0.35, fontSize: 13, color: "DDDDDD", margin: 0, isTextBox: true });
});
todo(s, "Agregar nombre, cargo, correo, teléfono y web de Mozart.", M, 6.25, 6.0, 0.45);

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})();
