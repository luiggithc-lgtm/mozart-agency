// Genera entregables/Propuesta-Mozart-Movitecnica.pptx
// Estilo: fondo carbón, fotos a sangre con capa de contraste, acento naranja, títulos en mayúsculas.
// Regla: las fotos van siempre a pantalla completa; el contraste del texto se logra con capas encima, nunca achicando la foto.
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/9e5afb32-f417-4b4e-8c86-5d1833191dc0_96639213-28e9-4b21-b321-4e5a36445164/pptx/scripts/apply_theme.js");

const OUT = path.join(__dirname, "../entregables/Propuesta-Mozart-Movitecnica.pptx");
const F = (f) => path.join(__dirname, "fotos", f);
const E = (f) => path.join(__dirname, "../entregables", f);

const C = { bg: "141414", card: "1F1F1F", line: "3A3A3A", orange: "F26B1D", white: "FFFFFF", grey: "C2C2C2", soft: "DDDDDD" };
const THEME = {
  name: "Mozart Industrial", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "141414", lt1: "FFFFFF", dk2: "1F1F1F", lt2: "F2F2F2", accent1: "F26B1D", accent2: "C2C2C2", accent3: "3A3A3A", accent4: "FFB27A", accent5: "8A8A8A", accent6: "5C5C5C", hlink: "F26B1D", folHlink: "C2C2C2" },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Propuesta Mozart para Movitécnica";
pres.author = "Mozart";
const W = 13.333, H = 7.5, M = 0.7;

pres.defineSlideMaster({
  title: "OSCURA", background: { color: C.bg },
  objects: [{ text: { text: "MOZART × MOVITÉCNICA", options: { x: M, y: 6.95, w: 5, h: 0.3, fontSize: 9, charSpacing: 3, color: "9A9A9A", margin: 0 } } }],
  slideNumber: { x: 12.2, y: 6.95, w: 0.45, h: 0.3, fontSize: 9, color: "9A9A9A", align: "right" },
});

// ---------- helpers ----------
const T = (s, t, o) => s.addText(t, { margin: 0, isTextBox: true, valign: "top", ...o });
const kicker = (s, t, x = M, y = 0.6) => T(s, t.toUpperCase(), { x, y, w: 9, h: 0.3, fontSize: 11, bold: true, charSpacing: 4, color: C.orange });
const title = (s, t, o = {}) => T(s, t.toUpperCase(), { x: M, y: 0.95, w: 8.5, h: 1.3, fontSize: 32, bold: true, color: C.white, ...o });
// Foto a pantalla completa + capa de contraste encima
const photo = (s, img, layer = "grad-left.png") => {
  s.addImage({ path: F(img), x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
  if (layer) s.addImage({ path: F(layer), x: 0, y: 0, w: W, h: H });
};
const corners = (s, x, y, w, h, len = 0.35) => {
  const L = (x1, y1, x2, y2) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001, line: { color: C.orange, width: 2.25 } });
  L(x, y, x + len, y); L(x, y, x, y + len); L(x + w - len, y, x + w, y); L(x + w, y, x + w, y + len);
  L(x, y + h, x + len, y + h); L(x, y + h - len, x, y + h); L(x + w - len, y + h, x + w, y + h); L(x + w, y + h - len, x + w, y + h);
};
const callout = (s, t, x, y, w, h, fontSize = 15) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.orange }, line: { color: C.orange } });
  T(s, t, { x: x + 0.25, y, w: w - 0.5, h, fontSize, bold: true, color: C.white, valign: "middle" });
};
const card = (s, x, y, w, h, glass = false) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: glass ? { color: "000000", transparency: 35 } : { color: C.card }, line: { color: glass ? "5A5A5A" : C.line, width: 0.75 } });
const num = (s, n, x, y, d = 0.5) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.orange }, line: { color: C.orange } });
  T(s, String(n), { x, y, w: d, h: d, fontSize: d < 0.5 ? 13 : 16, bold: true, color: C.white, align: "center", valign: "middle" });
};
const step = (s, n, head, text, x, y, w = 5.6) => {
  num(s, n, x, y);
  T(s, head.toUpperCase(), { x: x + 0.75, y: y - 0.04, w, h: 0.35, fontSize: 15, bold: true, color: C.white });
  T(s, text, { x: x + 0.75, y: y + 0.33, w, h: 0.5, fontSize: 13, color: C.grey });
};
const todo = (s, t, x, y, w, h = 0.45) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: "2A1A10" }, line: { color: C.orange, width: 1.25, dashType: "dash" } });
  T(s, [{ text: "LUIGGI · ", options: { bold: true } }, { text: t }], { x: x + 0.2, y, w: w - 0.4, h, fontSize: 11, color: "FFB27A", valign: "middle" });
};
const list = (s, items, o) => T(s, items.map((t, j) => ({ text: t, options: { bullet: { code: "25A0" }, breakLine: j < items.length - 1 } })), { fontSize: 13, color: C.grey, paraSpaceAfter: 6, ...o });
const add = (section) => pres.addSlide({ masterName: "OSCURA", sectionTitle: section });

// ---------- 1. Portada ----------
pres.addSection({ title: "Apertura" });
let s = add("Apertura");
photo(s, "faja-mina.jpg");
kicker(s, "Propuesta de aliado creativo · Octubre 2026", M, 1.2);
T(s, "INGENIERÍA\nQUE SE VE", { x: M, y: 2.0, w: 8, h: 2.6, fontSize: 64, bold: true, color: C.white });
T(s, "Web corporativa, contenido mensual para redes y producción audiovisual en campo para Movitécnica.", { x: M, y: 4.8, w: 5.8, h: 0.9, fontSize: 15, color: C.soft });
corners(s, 0.35, 0.35, W - 0.7, H - 0.7, 0.45);

// ---------- 2. Lo que entendimos ----------
pres.addSection({ title: "Contexto" });
s = add("Contexto");
photo(s, "polipasto.jpg");
kicker(s, "01 · Lo que entendimos");
title(s, "Mantener en movimiento operaciones que no pueden detenerse", { w: 6.6, fontSize: 30 });
[["46", "años en el mercado"], ["3", "países: Perú, Chile y Colombia"], ["~70%", "de sus ingresos viene de minería*"], ["4", "líneas: izaje, grúas, fajas y ascensores"]].forEach(([n, l], i) => {
  const x = M + (i % 2) * 3.25, y = 2.75 + Math.floor(i / 2) * 1.95;
  card(s, x, y, 3.0, 1.7, true);
  T(s, n, { x: x + 0.25, y: y + 0.15, w: 2.6, h: 0.85, fontSize: 40, bold: true, color: C.orange });
  T(s, l, { x: x + 0.25, y: y + 1.0, w: 2.6, h: 0.55, fontSize: 12, color: C.soft });
});
T(s, "* Según declaraciones públicas de la empresa.", { x: M, y: 6.62, w: 6, h: 0.25, fontSize: 9, color: "9A9A9A" });

// ---------- 3. El reto ----------
s = add("Contexto");
photo(s, "tecnico-polea.jpg", "grad-right.png");
kicker(s, "02 · El reto", 7.2);
title(s, "Nadie compra un puente grúa por impulso", { x: 7.2, w: 5.5, fontSize: 30 });
step(s, 1, "Operaciones", "Quiere ver capacidad técnica y soporte en faena.", 7.2, 2.75, 4.6);
step(s, 2, "Compras", "Necesita fichas técnicas y una forma rápida de cotizar.", 7.2, 3.95, 4.6);
step(s, 3, "Seguridad", "Exige evidencia de protocolos y trayectoria.", 7.2, 5.15, 4.6);

// ---------- 4. Tres frentes ----------
pres.addSection({ title: "Propuesta" });
s = add("Propuesta");
photo(s, "faja-mina.jpg", "dim.png");
kicker(s, "03 · Lo que ofrecemos");
title(s, "Tres servicios, contratables por separado o juntos", { w: 11 });
[["Web corporativa", "Rediseño completo de su web actual: más clara, ordenada por producto y pensada para recibir pedidos de cotización.", "Proyecto único"],
 ["Contenido mensual para redes", "Plan, diseño y publicación de reels, carruseles y posts para LinkedIn, Instagram y TikTok, todos los meses.", "Pago mensual"],
 ["Producción audiovisual en campo", "Grabación y fotografía en sus plantas y en faena: video institucional, casos de cliente y banco de fotos.", "Pago por jornada"]].forEach(([h, t, tag], i) => {
  const x = M + i * 4.1, y = 2.6, w = 3.75;
  card(s, x, y, w, 3.6, true);
  T(s, `0${i + 1}`, { x: x + 0.3, y: y + 0.3, w: 1, h: 0.5, fontSize: 28, bold: true, color: C.orange });
  T(s, h.toUpperCase(), { x: x + 0.3, y: y + 0.95, w: w - 0.6, h: 0.7, fontSize: 16, bold: true, color: C.white });
  T(s, t, { x: x + 0.3, y: y + 1.7, w: w - 0.6, h: 1.2, fontSize: 13, color: C.soft });
  T(s, tag.toUpperCase(), { x: x + 0.3, y: y + 3.1, w: w - 0.6, h: 0.3, fontSize: 10, bold: true, charSpacing: 2, color: C.orange });
});

// ---------- 5. Quiénes somos ----------
s = add("Propuesta");
photo(s, "cable-acero.jpg");
kicker(s, "04 · Quiénes somos");
title(s, "Oficio de red internacional, estructura ágil", { w: 6.5, fontSize: 30 });
step(s, 1, "Experiencia", "Equipo formado en TBWA y Publicis.", M, 2.75);
step(s, 2, "Reconocimientos", "Trabajos premiados en Cannes Lions y New York Festivals.", M, 3.85);
step(s, 3, "Trato directo", "Quien presenta la propuesta, la ejecuta.", M, 4.95);
todo(s, "Confirmar premio, año y pieza exactos según las Credenciales 2026.", M, 6.2, 6.4);

// ---------- 6. Casos ----------
s = add("Propuesta");
kicker(s, "05 · Experiencia relevante");
title(s, "Marcas que confiaron en nosotros", { w: 11 });
[["CLBS", "Web corporativa y rebranding"], ["ELECTRODATA", "Rebranding: experto, profesional, cercano"]].forEach(([h, t], i) => {
  const x = M + i * 6.1, w = 5.8;
  s.addShape(pres.shapes.RECTANGLE, { x, y: 2.2, w, h: 3.0, fill: { color: C.card }, line: { color: C.line, width: 0.75 } });
  corners(s, x + 0.15, 2.35, w - 0.3, 2.7, 0.3);
  T(s, "Imagen del caso", { x, y: 3.5, w, h: 0.4, fontSize: 12, color: "6B6B6B", align: "center" });
  T(s, h, { x, y: 5.4, w, h: 0.4, fontSize: 18, bold: true, color: C.white });
  T(s, t, { x, y: 5.8, w, h: 0.35, fontSize: 13, color: C.grey });
});
T(s, [{ text: "También: ", options: { color: C.grey } }, { text: "BTG PACTUAL · SERRA · LYNCH", options: { bold: true, color: C.white } }], { x: M, y: 6.35, w: 6, h: 0.35, fontSize: 12 });
todo(s, "Insertar imágenes de CLBS y Electrodata.", 7.6, 6.3, 5.0);

// ---------- 7. Concepto ----------
pres.addSection({ title: "Concepto" });
s = add("Concepto");
photo(s, "puente-grua.jpg");
kicker(s, "06 · Concepto", M, 1.3);
T(s, "INGENIERÍA\nQUE SE VE", { x: M, y: 1.8, w: 7, h: 2.4, fontSize: 60, bold: true, color: C.white });
callout(s, "Su trabajo ocurre en faena, de noche y en paradas de planta. Proponemos mostrarlo.", M, 4.45, 5.6, 0.95);
["Mostrar el proceso", "Hablar en datos", "La seguridad como valor"].forEach((t, i) => {
  card(s, M + i * 2.05, 5.7, 1.9, 0.55, true);
  T(s, t, { x: M + i * 2.05, y: 5.7, w: 1.9, h: 0.55, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
});

// ---------- 8. Web: qué proponemos ----------
pres.addSection({ title: "Web corporativa" });
s = add("Web corporativa");
photo(s, "puente-grua.jpg", "dim.png");
kicker(s, "Servicio 01 · Web corporativa");
title(s, "Su web actual, rediseñada para generar cotizaciones", { w: 11 });
[["Mismo contenido, mejor ordenado", "Partimos de su web de hoy: mismas secciones, productos y servicios, con una navegación más simple."],
 ["Un botón de cotización en cada producto", "Cada ficha tiene el botón \"Solicitar cotización\". El formulario ya llega con el nombre del equipo, así ventas sabe qué pide el cliente sin preguntar."],
 ["Una versión por país", "Perú, Chile y Colombia con sus propios contactos, sedes y productos disponibles."]].forEach(([h, t], i) => {
  const x = M + i * 4.1, y = 2.6, w = 3.75;
  card(s, x, y, w, 3.4, true);
  num(s, i + 1, x + 0.3, y + 0.3);
  T(s, h.toUpperCase(), { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: 0.7, fontSize: 15, bold: true, color: C.white });
  T(s, t, { x: x + 0.3, y: y + 1.75, w: w - 0.6, h: 1.5, fontSize: 13, color: C.soft });
});

// ---------- 9. Web: mockup a tamaño completo ----------
s = add("Web corporativa");
s.addImage({ path: E("muestra-home-movitecnica.png"), x: (W - 9.92) / 2, y: 0.25, w: 9.92, h: 6.2 });
callout(s, "Muestra referencial de la nueva home: botón de cotización en cada producto y selector de país", (W - 9.92) / 2, 6.5, 9.92, 0.4, 12);

// ---------- 10. Web: alcance según su web actual ----------
s = add("Web corporativa");
kicker(s, "Servicio 01 · Alcance");
title(s, "Páginas incluidas, según su web actual", { w: 11.5, h: 0.6 });
const cols = [
  ["Páginas principales", ["Home", "Nosotros", "Contacto con sedes (Lima, Arequipa, Chiclayo, Chile, Colombia)", "Formulario de cotización", "Verifica tu equipo / certificado"]],
  ["Productos", ["Página general de productos", "Categorías: tecles y polipastos, grúas (puente grúa, pórtico, pescantes), fajas transportadoras, ascensores industriales, carretillas hidráulicas, eslingas y cables", "Ficha de producto con botón de cotización: carga de hasta ___ productos"]],
  ["Servicios y blog", ["Página general de servicios", "5 páginas: polipastos, fajas, ascensores, entrenamiento y alquiler de andamios", "Blog: listado y artículo, con migración de las notas actuales"]],
];
cols.forEach(([h, items], i) => {
  const x = M + i * 4.1, y = 1.9, w = 3.75;
  card(s, x, y, w, 4.15);
  T(s, h.toUpperCase(), { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 0.35, fontSize: 14, bold: true, color: C.orange });
  list(s, items, { x: x + 0.3, y: y + 0.8, w: w - 0.6, h: 3.2, fontSize: 12 });
});
todo(s, "Contar los productos de su web actual para fijar el número de fichas incluidas. Validar categorías con el cliente.", M, 6.25, 11.9);

// ---------- 11. Web: qué incluye el trabajo ----------
s = add("Web corporativa");
kicker(s, "Servicio 01 · Qué incluye el trabajo");
title(s, "De la auditoría al lanzamiento", { w: 11 });
[["Auditoría", "Revisión de su web actual: qué funciona y qué no."], ["Arquitectura", "Menú y orden de páginas, pensado para quien cotiza."], ["Wireframes UX", "Esquema de cada plantilla antes de diseñar."],
 ["Diseño UI", "Diseño visual en computadora y celular."], ["Desarrollo", "Web autoadministrable, SEO técnico y migración de contenidos."], ["Lanzamiento", "Publicación, pruebas y capacitación a su equipo."]].forEach(([h, t], i) => {
  const x = M + (i % 3) * 4.1, y = 2.2 + Math.floor(i / 3) * 1.85;
  card(s, x, y, 3.75, 1.6);
  num(s, i + 1, x + 0.25, y + 0.25, 0.45);
  T(s, h.toUpperCase(), { x: x + 0.85, y: y + 0.28, w: 2.7, h: 0.4, fontSize: 14, bold: true, color: C.white });
  T(s, t, { x: x + 0.25, y: y + 0.85, w: 3.3, h: 0.6, fontSize: 12, color: C.grey });
});
todo(s, "Definir rondas de cambios incluidas y si hosting y dominio los mantiene Movitécnica.", M, 6.1, 9.0);

// ---------- 12. Contenido mensual para redes ----------
pres.addSection({ title: "Contenido mensual" });
s = add("Contenido mensual");
photo(s, "tecnico-polea.jpg", "grad-right.png");
kicker(s, "Servicio 02 · Contenido mensual para redes", 6.9);
title(s, "Qué recibe Movitécnica cada mes", { x: 6.9, w: 5.8, fontSize: 30, h: 1.0 });
[["4", "reels verticales con subtítulos"], ["4", "carruseles técnicos"], ["4", "publicaciones de imagen"]].forEach(([n, l], i) => {
  const y = 2.2 + i * 0.9;
  card(s, 6.9, y, 5.7, 0.75, true);
  T(s, n, { x: 7.15, y, w: 0.8, h: 0.75, fontSize: 30, bold: true, color: C.orange, valign: "middle" });
  T(s, l.toUpperCase(), { x: 7.95, y, w: 4.5, h: 0.75, fontSize: 14, bold: true, color: C.white, valign: "middle" });
});
list(s, ["Plan de contenidos y textos para aprobar", "Publicación en LinkedIn, Instagram y TikTok", "Reporte mensual de resultados"], { x: 6.9, y: 4.95, w: 5.7, h: 0.95, color: C.soft });
todo(s, "Ajustar cantidades. El material de video sale de las jornadas en campo o de material del cliente.", 6.9, 6.05, 5.7, 0.6);

// ---------- 13. Storyboard a tamaño completo ----------
s = add("Contenido mensual");
s.addImage({ path: E("muestra-storyboard-reel.png"), x: 0, y: 0, w: W, h: H });

// ---------- 14. Producción audiovisual en campo ----------
pres.addSection({ title: "Audiovisual en campo" });
s = add("Audiovisual en campo");
photo(s, "camarografo.jpg");
kicker(s, "Servicio 03 · Producción audiovisual en campo", M, 0.6);
title(s, "Qué incluye una jornada de rodaje", { w: 6.5, fontSize: 30 });
[["Equipo", "Director, camarógrafo y fotógrafo en su planta o faena."], ["Duración", "Hasta 8 horas de rodaje en una sede."], ["Entregables", "1 video editado de hasta 90 s, 3 cortes verticales para redes y 20 fotos editadas."]].forEach(([h, t], i) => {
  const y = 2.3 + i * 1.15;
  card(s, M, y, 6.0, 0.95, true);
  T(s, h.toUpperCase(), { x: M + 0.25, y: y + 0.12, w: 5.5, h: 0.35, fontSize: 13, bold: true, color: C.orange });
  T(s, t, { x: M + 0.25, y: y + 0.47, w: 5.5, h: 0.45, fontSize: 13, color: C.soft });
});
T(s, "Opcionales: video institucional de 2 a 3 minutos, casos de cliente y tomas con dron.", { x: M, y: 5.8, w: 6.0, h: 0.4, fontSize: 12, color: C.soft });
todo(s, "Ajustar equipo, horas y entregables por jornada.", M, 6.3, 6.0, 0.4);

// ---------- 15. Seguridad ----------
s = add("Audiovisual en campo");
photo(s, "charla-seguridad.jpg");
kicker(s, "Protocolo de seguridad en campo");
title(s, "Primero la seguridad, después la cámara", { w: 6.2, fontSize: 30 });
step(s, 1, "Antes", "Inducción y plan de rodaje aprobado por la unidad.", M, 2.7);
step(s, 2, "Durante", "EPP completo y acompañamiento de Movitécnica.", M, 3.75);
step(s, 3, "Después", "Revisión del material antes de publicar.", M, 4.8);
todo(s, "Contratar SCTR para cada persona del equipo antes de entrar a faena.", M, 5.85, 6.1, 0.42);
todo(s, "Conseguir piloto de dron con registro vigente y seguro.", M, 6.37, 6.1, 0.42);

// ---------- 16. Cómo trabajamos ----------
pres.addSection({ title: "Cierre" });
s = add("Cierre");
kicker(s, "Cómo trabajamos");
title(s, "Un proceso claro en cuatro fases", { w: 11 });
s.addShape(pres.shapes.LINE, { x: M + 0.3, y: 3.15, w: 11.3, h: 0, line: { color: C.line, width: 1.5 } });
[["Inmersión", "Marca, catálogo y prioridades por país."], ["Estrategia y diseño", "Arquitectura web, línea visual y plan de contenidos."], ["Producción", "Desarrollo web y jornadas de rodaje."], ["Lanzamiento", "Publicación, medición y ajustes mensuales."]].forEach(([h, t], i) => {
  const x = M + i * 3.0;
  num(s, i + 1, x + 0.05, 2.9);
  T(s, h.toUpperCase(), { x, y: 3.7, w: 2.7, h: 0.4, fontSize: 15, bold: true, color: C.white });
  T(s, t, { x, y: 4.15, w: 2.6, h: 0.8, fontSize: 13, color: C.grey });
});
todo(s, "Definir plazos por fase (en semanas).", M, 5.8, 5.0);

// ---------- 17. Inversión ----------
s = add("Cierre");
kicker(s, "Inversión");
title(s, "Inversión por servicio", { w: 11, h: 0.6 });
[["Servicio 01 · Proyecto único", "Web corporativa", "S/ ____________", ["Rediseño de su web actual", "Plantillas y páginas del alcance", "Fichas con botón de cotización", "Versiones Perú, Chile y Colombia", "SEO técnico y capacitación"]],
 ["Servicio 02 · Pago mensual", "Contenido mensual para redes", "S/ ________ /mes", ["4 reels verticales", "4 carruseles técnicos", "4 publicaciones de imagen", "Plan, textos y publicación", "Reporte mensual"]],
 ["Servicio 03 · Por jornada", "Audiovisual en campo", "S/ ______ /jornada", ["Hasta 8 horas de rodaje", "Director, camarógrafo y fotógrafo", "1 video de hasta 90 s", "3 cortes verticales", "20 fotos editadas"]]].forEach(([k, h, p, items], i) => {
  const x = M + i * 4.1, y = 1.85, w = 3.75;
  card(s, x, y, w, 4.55);
  T(s, k.toUpperCase(), { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 0.3, fontSize: 10, bold: true, charSpacing: 2, color: C.orange });
  T(s, h.toUpperCase(), { x: x + 0.3, y: y + 0.65, w: w - 0.6, h: 0.65, fontSize: 15, bold: true, color: C.white });
  T(s, p, { x: x + 0.3, y: y + 1.35, w: w - 0.6, h: 0.5, fontSize: 22, bold: true, color: C.white });
  list(s, items, { x: x + 0.3, y: y + 2.1, w: w - 0.6, h: 2.3 });
});
T(s, "Montos en soles. IGV y viáticos a faena: ____________", { x: M, y: 6.6, w: 8, h: 0.3, fontSize: 11, color: "9A9A9A" });

// ---------- 18. Cierre ----------
s = add("Cierre");
photo(s, "faja-mina.jpg", "grad-full.png");
kicker(s, "Próximos pasos", M, 1.2);
T(s, "CONVERSEMOS", { x: M, y: 1.65, w: 8, h: 1.2, fontSize: 60, bold: true, color: C.white });
step(s, 1, "Reunión de 30 minutos", "Prioridades y alcance por país.", M, 3.3);
step(s, 2, "Propuesta ajustada", "Alcance y montos finales.", M, 4.25);
step(s, 3, "Arranque", "Inmersión y primera jornada en campo.", M, 5.2);
todo(s, "Agregar nombre, cargo, correo, teléfono y web de Mozart.", M, 6.25, 6.0);

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})();
