const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("C:/Users/Nico/.claude/skills/synced/232049fb-98e4-4c3d-a47e-9a56979acece_05762fec-e168-47d2-9836-a44ed0fdebaa/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "presentacion.pptx";
const FIG = path.join(__dirname, "fig");
const BLOQUES = "D:/CEIA/4to Bimestre/GdP/Planificacion-AGILE/Figuras/diagBloques-proyecto.png";

const THEME = {
  name: "Plan TF CEIA",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2D27", lt1: "FFFFFF", dk2: "1B4332", lt2: "EEF6F0",
    accent1: "2D6A4F", accent2: "E8590C", accent3: "95D5B2", accent4: "FFE8CC",
    accent5: "52796F", accent6: "B35C00", hlink: "2D6A4F", folHlink: "52796F",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Presentación del plan de Trabajo Final";
pres.author = "Ing. Nicolás Francisco Moguilner Reh";
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: "FFFFFF" },
  objects: [
    { image: { path: path.join(FIG, "logo.png"), x: 0.5, y: 5.17, w: 0.9, h: 0.3 } },
    { text: { text: "Plan de Trabajo Final · CEIA · FIUBA", options: { x: 1.5, y: 5.17, w: 5, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, valign: "middle" } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 9, h: 0.75, fontSize: 32, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.0, y: 5.17, w: 0.5, h: 0.3, fontSize: 12, color: C.text2, align: "right", bold: true },
});
pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: "FFFFFF" },
  objects: [],
  slideNumber: { x: 5.2, y: 5.17, w: 0.5, h: 0.3, fontSize: 12, color: C.text2, align: "right", bold: true },
});
pres.defineSlideMaster({
  title: "CIERRE",
  background: { color: THEME.colors.dk2 },
  objects: [],
  slideNumber: { x: 9.0, y: 5.17, w: 0.5, h: 0.3, fontSize: 12, color: C.background1, align: "right", bold: true },
});

const card = (s, x, y, w, h, fill, name) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
const tx = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top" }, o));

// ---------- 1. Portada ----------
pres.addSection({ title: "Carátula" });
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Carátula" });
  s.addImage({ path: path.join(FIG, "oferta_coto.jpg"), x: 6.0, y: 0, w: 4.0, h: 5.625, altText: "Cartel de oferta 3x2 en una góndola de Coto", objectName: "Foto de supermercado" });
  tx(s, "Foto: r/argentina, Reddit", { x: 0.5, y: 5.2, w: 3.5, h: 0.22, fontSize: 9, color: C.accent5 });
  s.addImage({ path: path.join(FIG, "logo.png"), x: 0.5, y: 0.35, w: 2.1, h: 0.7, objectName: "Logo FIUBA" });
  tx(s, "PRESENTACIÓN DEL PLAN DE TRABAJO FINAL", { x: 0.5, y: 1.25, w: 5.2, h: 0.3, fontSize: 13, bold: true, color: C.accent2, charSpacing: 1 });
  tx(s, "Carrera de Especialización en Inteligencia Artificial", { x: 0.5, y: 1.55, w: 5.2, h: 0.3, fontSize: 14, color: C.accent5 });
  tx(s, [
    { text: "Sistema inteligente de optimización de compras de supermercado", options: { fontSize: 24, bold: true, color: C.text2, breakLine: true } },
    { text: "con pronóstico de precios y detección de ofertas engañosas", options: { fontSize: 15, color: C.accent1 } },
  ], { x: 0.5, y: 1.95, w: 5.35, h: 1.7, fontFace: THEME.headFontFace, valign: "middle", paraSpaceAfter: 4 });
  card(s, 0.5, 3.8, 5.2, 1.2, THEME.colors.lt2, "Datos del trabajo");
  tx(s, [
    { text: "Alumno: ", options: { bold: true } }, { text: "Ing. Nicolás Francisco Moguilner Reh", options: { breakLine: true } },
    { text: "Director: ", options: { bold: true } }, { text: "Dr. Ing. Facundo Adrián Lucianna", options: { breakLine: true } },
    { text: "Contexto: ", options: { bold: true } }, { text: "FIUBA, noviembre de 2026" },
  ], { x: 0.7, y: 3.9, w: 4.9, h: 1.0, fontSize: 15, color: C.text1, valign: "middle", paraSpaceAfter: 2 });
  s.addNotes("Presentación del plan del Trabajo Final de la CEIA. Título, alumno y director.");
}

// ---------- 2. El problema ----------
pres.addSection({ title: "Introducción" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Introducción" });
  s.addText("¿Dónde, cuándo y con qué pagar la compra?", { placeholder: "title" });
  const items = [
    ["Precios que cambian", "Las cadenas remarcan cada pocas semanas"],
    ["Promociones difíciles de comparar", "Cada banco y billetera tiene días, topes, segmentos y exclusiones"],
    ["Ofertas engañosas", "El precio sube antes de la promoción para luego “bajar”"],
  ];
  items.forEach(([h, d], i) => {
    const y = 1.2 + i * 1.27;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `Número ${i + 1}` });
    tx(s, String(i + 1), { x: 0.5, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
    tx(s, h, { x: 1.2, y: y, w: 3.8, h: 0.4, fontSize: 19, bold: true, color: C.text2 });
    tx(s, d, { x: 1.2, y: y + 0.42, w: 3.8, h: 0.8, fontSize: 15, color: C.text1 });
  });
  const th = 3.7, tw = th * 1640 / 3300;
  s.addImage({ path: path.join(FIG, "ticket.jpg"), x: 5.2, y: 1.15, w: tw, h: th, altText: "Ticket de Coto del 16/11/2024 con un descuento de Mercado Pago del 15 % que rindió $ 1.172,56 sobre $ 33.372", objectName: "Foto del ticket" });
  tx(s, "Fuente: publicación en Threads", { x: 5.2, y: 4.87, w: 2.1, h: 0.22, fontSize: 10, color: C.accent5 });
  card(s, 7.35, 1.15, 2.15, 3.7, THEME.colors.lt2, "Ejemplo de ticket");
  tx(s, "Coto, 16/11/2024", { x: 7.5, y: 1.28, w: 1.9, h: 0.3, fontSize: 13, bold: true, color: C.accent5 });
  tx(s, "15 %", { x: 7.5, y: 1.65, w: 1.9, h: 0.7, fontSize: 40, bold: true, color: C.accent5 });
  tx(s, "descuento anunciado de Mercado Pago", { x: 7.5, y: 2.35, w: 1.9, h: 0.55, fontSize: 13, color: C.text1 });
  tx(s, "3,5 %", { x: 7.5, y: 3.05, w: 1.9, h: 0.7, fontSize: 40, bold: true, color: C.accent2 });
  tx(s, "descuento efectivo por exclusiones", { x: 7.5, y: 3.75, w: 1.9, h: 0.55, fontSize: 13, color: C.text1 });
  s.addNotes("El problema: el consumidor no puede comparar fácilmente precios que cambian, promociones con muchas condiciones y ofertas que no son reales. Ejemplo: un 15 % anunciado terminó siendo un 3,5 % efectivo.");
}

// ---------- 3. La solución ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Introducción" });
  s.addText("La solución propuesta", { placeholder: "title" });
  const w = 9, h = w * 847 / 2000;
  s.addImage({ path: BLOQUES, x: 0.5, y: 1.15, w, h, altText: "Diagrama en bloques del sistema propuesto", objectName: "Diagrama en bloques" });
  s.addNotes("Sobre los datos abiertos del SEPA, el sistema recomienda sucursal, día y medio de pago, y alerta sobre ofertas engañosas. Diagrama en bloques: fuentes de datos, ingesta, unificación de productos, pronóstico, detección de ofertas engañosas, optimizador y aplicación web. Los bloques grises punteados son opcionales.");
}

// ---------- 4. Los datos ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Introducción" });
  s.addText("Los datos disponibles", { placeholder: "title" });
  const stats = [
    ["3.100", "sucursales informan a diario sus precios al SEPA (datos abiertos)"],
    ["40 GB", "de histórico de precios desde agosto de 2024"],
    ["84 %", "de los días con datos: 687 de 818, con huecos conocidos"],
    ["480", "promociones bancarias y de billeteras reconstruidas de 80 fuentes"],
  ];
  stats.forEach(([n, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.85;
    card(s, x, y, 4.4, 1.65, THEME.colors.lt2, `Dato ${i + 1}`);
    tx(s, n, { x: x + 0.25, y: y + 0.18, w: 3.9, h: 0.75, fontSize: 40, bold: true, color: i === 2 ? C.accent2 : C.accent1 });
    tx(s, d, { x: x + 0.25, y: y + 0.95, w: 3.9, h: 0.6, fontSize: 15, color: C.text1 });
  });
  s.addNotes("El portal oficial solo conserva siete días, por eso se incorpora el histórico de preciazo. Hay huecos: faltan 105 días entre marzo y junio de 2026. Las promociones se reconstruyen de notas periodísticas.");
}

// ---------- 5. Interesados ----------
pres.addSection({ title: "Plan" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Los interesados", { placeholder: "title" });
  const st = [
    ["Cliente", "Dr. Ing. Facundo A. Lucianna", "Valida el plan y aprueba los resultados"],
    ["Impulsor", "Federico G. Zacchigna", "Coordinación académica de IA"],
    ["Responsable", "Ing. Nicolás F. Moguilner Reh", "Alumno, desarrollo individual"],
    ["Orientador", "Dr. Ing. Facundo A. Lucianna", "Director del Trabajo Final"],
    ["Opositores", "Cadenas de supermercados", "Podrían objetar el análisis de ofertas"],
    ["Usuario final", "Consumidores", "Compran en supermercados con varios medios de pago"],
  ];
  st.forEach(([r, n, d], i) => {
    const x = 0.5 + (i % 3) * 3.07, y = 1.25 + Math.floor(i / 3) * 1.95;
    card(s, x, y, 2.86, 1.78, THEME.colors.lt2, `Interesado ${r}`);
    tx(s, r, { x: x + 0.2, y: y + 0.15, w: 2.5, h: 0.45, fontSize: 22, bold: true, color: i === 4 ? C.accent2 : C.accent1 });
    tx(s, n, { x: x + 0.2, y: y + 0.62, w: 2.5, h: 0.4, fontSize: 15, bold: true, color: C.text1 });
    tx(s, d, { x: x + 0.2, y: y + 1.02, w: 2.5, h: 0.65, fontSize: 14, color: C.text1 });
  });
  s.addNotes("El director cumple dos roles: cliente, que aprueba los resultados, y orientador. El impulsor sugirió la consulta de precios de comercios fuera de las cadenas.");
}

// ---------- 6. Propósito ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("El propósito", { placeholder: "title" });
  card(s, 0.5, 1.3, 9, 2.6, THEME.colors.lt2, "Propósito");
  tx(s, "Ayudar a los consumidores a pagar menos por su compra de supermercado, recomendándoles dónde, cuándo y con qué medio de pago comprar, y protegiéndolos de ofertas engañosas", { x: 0.85, y: 1.45, w: 8.3, h: 2.3, fontSize: 26, bold: true, color: C.text2, valign: "middle" });
  tx(s, [
    { text: "Valor medible: ", options: { bold: true, color: C.accent2 } },
    { text: "el ahorro en pesos frente a la forma habitual de comprar (misma sucursal, el día de siempre, cualquier medio de pago)" },
  ], { x: 0.5, y: 4.1, w: 9, h: 0.8, fontSize: 18, color: C.text1 });
  s.addNotes("El propósito se mide en pesos ahorrados, no solo en métricas técnicas.");
}

// ---------- 7. Alcance ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("El alcance", { placeholder: "title" });
  card(s, 0.5, 1.2, 5.6, 3.75, THEME.colors.lt2, "Incluye");
  tx(s, "Incluye", { x: 0.75, y: 1.32, w: 5.1, h: 0.45, fontSize: 22, bold: true, color: C.accent1 });
  const inc = [
    "Descarga diaria del SEPA e histórico desde 2024",
    "Coto, Carrefour, ChangoMás y DIA en Paraná",
    "Unificación de productos y precio por unidad",
    "Pronóstico de precios y de remarcación",
    "Detección acotada de ofertas engañosas",
    "Optimizador, API y aplicación web de demostración",
    "Evaluación del ahorro por backtest",
  ];
  tx(s, inc.map((t, i) => ({ text: t, options: { bullet: { indent: 15 }, breakLine: i < inc.length - 1 } })), { x: 0.75, y: 1.85, w: 5.2, h: 3.0, fontSize: 17, color: C.text1, paraSpaceAfter: 5 });
  card(s, 6.3, 1.2, 3.2, 3.75, THEME.colors.accent4, "No incluye");
  tx(s, "No incluye", { x: 6.55, y: 1.32, w: 2.8, h: 0.45, fontSize: 22, bold: true, color: C.accent6 });
  const exc = ["Puesta en producción", "Productos sin código de barras", "Comercios que no informan al SEPA", "Comercio electrónico"];
  tx(s, exc.map((t, i) => ({ text: t, options: { bullet: { indent: 15 }, breakLine: i < exc.length - 1 } })), { x: 6.55, y: 1.85, w: 2.8, h: 3.0, fontSize: 17, color: C.text1, paraSpaceAfter: 5 });
  s.addNotes("Opcionales si el cronograma lo permite: más cadenas, lenguaje natural, crawlers de promociones, sustitutos, costo de traslado y lectura de tickets.");
}

// ---------- 8. Requerimientos ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Los requerimientos: Product Backlog", { placeholder: "title" });
  const ep = [
    ["Datos", "34 SP", ["HU1 Descarga diaria", "HU2 Validación", "HU3 Promociones"]],
    ["Unificación", "55 SP", ["HU4 Mismo producto en todas las cadenas", "HU5 Equivalentes"]],
    ["Pronóstico y ofertas", "76 SP", ["HU6 Precio a 14 días", "HU7 Alertas de ofertas engañosas"]],
    ["Recomendación", "81 SP", ["HU8 Plan de compra", "HU9 Precio externo", "HU10 App web", "HU11 Ahorro medido"]],
  ];
  ep.forEach(([t, sp, hus], i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.2, 2.15, 3.2, THEME.colors.lt2, `Épica ${i + 1}`);
    tx(s, `Épica ${i + 1}`, { x: x + 0.18, y: 1.3, w: 1.85, h: 0.3, fontSize: 13, color: C.accent5 });
    tx(s, t, { x: x + 0.18, y: 1.58, w: 1.85, h: 0.7, fontSize: 19, bold: true, color: C.text2 });
    tx(s, sp, { x: x + 0.18, y: 2.25, w: 1.85, h: 0.4, fontSize: 18, bold: true, color: C.accent2 });
    tx(s, hus.map((h, j) => ({ text: h, options: { breakLine: j < hus.length - 1 } })), { x: x + 0.18, y: 2.72, w: 1.85, h: 1.6, fontSize: 14, color: C.text1, paraSpaceAfter: 4 });
  });
  tx(s, [
    { text: "246 SP en el núcleo", options: { bold: true, color: C.text2 } },
    { text: "  ·  1 SP ≈ 2 h  ·  opcionales: HU12 lenguaje natural y HU13 crawlers de promociones" },
  ], { x: 0.5, y: 4.55, w: 9, h: 0.4, fontSize: 15, color: C.text1 });
  s.addNotes("Trece historias de usuario en cuatro épicas. Las de mayor peso son el pronóstico (HU6, 55 SP) y la unificación de equivalentes (HU5, 34 SP).");
}

// ---------- 9. Criterios de aceptación ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Criterios de aceptación de HU", { placeholder: "title" });
  const q = [
    ["HU1", "≥ 95 %", "de los días publicados quedan descargados"],
    ["HU5", "F1 ≥ 0,80", "al detectar productos equivalentes"],
    ["HU6", "MAPE", "menor que la línea de base, con intervalos del 80 %"],
    ["HU7", "≥ 0,80", "de precisión en las alertas de ofertas engañosas"],
    ["HU8", "< 10 s", "por plan de compra, igual al óptimo exhaustivo"],
    ["HU11", "200 listas", "simuladas para medir el ahorro con IC del 95 %"],
  ];
  q.forEach(([hu, n, d], i) => {
    const x = 0.5 + (i % 3) * 3.07, y = 1.2 + Math.floor(i / 3) * 1.6;
    card(s, x, y, 2.86, 1.45, THEME.colors.lt2, `Criterio ${hu}`);
    tx(s, n, { x: x + 0.2, y: y + 0.12, w: 1.9, h: 0.55, fontSize: 26, bold: true, color: C.accent1 });
    tx(s, hu, { x: x + 2.0, y: y + 0.18, w: 0.66, h: 0.35, fontSize: 14, bold: true, color: C.accent2, align: "right" });
    tx(s, d, { x: x + 0.2, y: y + 0.72, w: 2.5, h: 0.65, fontSize: 14, color: C.text1 });
  });
  tx(s, "Una historia está completa cuando cumple todos sus criterios", { x: 0.5, y: 4.45, w: 9, h: 0.45, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  s.addNotes("Cada historia tiene criterios medibles; estos son los más representativos. Todos los modelos se comparan contra su línea de base y la evaluación es reproducible con un único script.");
}

// ---------- 10. CRISP-DM ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Fases de CRISP-DM", { placeholder: "title" });
  const f = [
    ["Comprensión del negocio", "Pagar lo mínimo; métrica: ahorro en pesos"],
    ["Comprensión de los datos", "SEPA, preciazo y promociones, con huecos medidos"],
    ["Preparación de los datos", "Validación, unificación y partición temporal"],
    ["Modelado", "Pronóstico por cuantiles, clasificadores y optimización entera"],
    ["Evaluación", "Métricas por módulo y backtest del ahorro"],
    ["Despliegue", "Demo con FastAPI, Docker, PostgreSQL y MLflow"],
  ];
  f.forEach(([t, d], i) => {
    const x = 0.5 + (i % 3) * 3.07, y = 1.2 + Math.floor(i / 3) * 1.85;
    card(s, x, y, 2.86, 1.65, THEME.colors.lt2, `Fase ${i + 1}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.18, w: 0.45, h: 0.45, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `Número fase ${i + 1}` });
    tx(s, String(i + 1), { x: x + 0.2, y: y + 0.18, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    tx(s, t, { x: x + 0.78, y: y + 0.12, w: 1.95, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    tx(s, d, { x: x + 0.2, y: y + 0.82, w: 2.5, h: 0.75, fontSize: 14, color: C.text1 });
  });
  s.addNotes("La partición es temporal: entrenamiento de agosto de 2024 a noviembre de 2025, validación hasta marzo de 2026 y prueba de julio a septiembre de 2026. El modelado combina pronóstico por cuantiles, clasificación de remarcación y de pares, Isolation Forest y programación entera mixta.");
}

// ---------- 10. Gantt ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Diagrama de Gantt", { placeholder: "title" });
  const h = 3.95, w = h * 2561 / 1516;
  s.addImage({ path: path.join(FIG, "gantt.png"), x: 0.5, y: 1.1, w, h, altText: "Diagrama de Gantt de los sprints", objectName: "Diagrama de Gantt" });
  const st = [["13", "sprints de 3 semanas"], ["600 h", "492 técnicas y 108 no técnicas"], ["20/06/2027", "fin del desarrollo técnico"], ["16/08/2027", "defensa pública"]];
  st.forEach(([n, d], i) => {
    const y = 1.1 + i * 1.0;
    tx(s, n, { x: 7.35, y, w: 2.2, h: 0.45, fontSize: 24, bold: true, color: i === 3 ? C.accent2 : C.accent1 });
    tx(s, d, { x: 7.35, y: y + 0.45, w: 2.2, h: 0.45, fontSize: 13, color: C.text1 });
  });
  s.addNotes("Sprint 0 de planificación durante la cursada. Sprints técnicos hasta junio de 2027; la memoria se redacta durante los talleres A y B, de abril a agosto de 2027. Costo estimado: USD 12.600.");
}

// ---------- 11. Riesgos ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Gestión de riesgos", { placeholder: "title" });
  const rows = [
    ["Interrupción de la descarga diaria del SEPA", 8, 4, 32, 5, 2, 10],
    ["Cobertura insuficiente del histórico de promociones", 7, 7, 49, 5, 5, 25],
    ["El pronóstico no supera a la línea de base", 6, 6, 36, 4, 5, 20],
    ["Falta de disponibilidad horaria del autor", 7, 6, 42, 5, 4, 20],
    ["Errores en la unificación de productos", 6, 5, 30, "-", "-", "-"],
    ["Pérdida de los datos históricos", 9, 3, 27, "-", "-", "-"],
    ["Demora o costos de la infraestructura en la nube", 5, 5, 25, "-", "-", "-"],
  ];
  const head = ["Riesgo", "S", "O", "RPN", "S*", "O*", "RPN*"].map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.colors.dk2 }, align: i ? "center" : "left" } }));
  const body = rows.map((r, k) => r.map((v, i) => {
    const hi = i === 3 && v > 30;
    return { text: i === 0 ? `${k + 1}. ${v}` : String(v), options: { align: i ? "center" : "left", bold: hi, color: hi ? "FFFFFF" : THEME.colors.dk1, fill: { color: hi ? THEME.colors.accent2 : (k % 2 ? "FFFFFF" : THEME.colors.lt2) } } };
  }));
  s.addTable([head, ...body], { x: 0.5, y: 1.15, w: 9, colW: [5.1, 0.6, 0.6, 0.75, 0.6, 0.6, 0.75], fontSize: 15, rowH: 0.42, border: { type: "solid", pt: 0.75, color: "C9DCCF" }, valign: "middle", margin: [0.02, 0.08, 0.02, 0.08] });
  tx(s, [
    { text: "RPN = S × O. ", options: { bold: true } },
    { text: "Se mitigan los riesgos con RPN > 30 (en naranja). S*, O* y RPN*: valores después de mitigar" },
  ], { x: 0.5, y: 4.62, w: 9, h: 0.35, fontSize: 14, color: C.text1 });
  s.addNotes("Riesgo principal: la cobertura del histórico de promociones, porque afecta la métrica central. Se mitiga evaluando por separado según la cobertura y con un análisis de sensibilidad de los topes.");
}

// ---------- 13. Retrospectiva ----------
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Plan" });
  s.addText("Sprint Retrospective", { placeholder: "title" });
  const W = THEME.colors;
  const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: W.dk2 }, align: "left" } });
  const ax = (t) => ({ text: t, options: { bold: true, color: W.dk2, fill: { color: "D8F3DC" } } });
  const cell = (t, k) => ({ text: t, options: { color: W.dk1, fill: { color: k % 2 ? "FFFFFF" : W.lt2 } } });
  const rows = [
    ["Hacer más", "Revisar a diario el inventario de fechas", "Comparar cada modelo con la línea de base"],
    ["Hacer menos", "Procesar datos de todo el país", "Ajustar hiperparámetros sin registro"],
    ["Mantener", "Guardar cada archivo original con su hash", "Registrar los experimentos en MLflow"],
    ["Empezar a hacer", "Alertas automáticas ante fallas de descarga", "Analizar errores por cadena y categoría"],
    ["Dejar de hacer", "Descargas manuales fuera del proceso", "Mirar el período de prueba al desarrollar"],
  ];
  const tasks = [
    { text: "Tareas", options: { bold: true, color: W.dk2, fill: { color: "D8F3DC" } } },
    { text: "Descarga diaria del SEPA, histórico de preciazo y esquema de promociones", options: { italic: false, color: W.accent5, fill: { color: "FFFFFF" } } },
    { text: "Modelo de pronóstico por cuantiles y evaluación de MAPE y cobertura", options: { color: W.accent5, fill: { color: "FFFFFF" } } },
  ];
  s.addTable([
    [hd("Eje"), hd("Sprint 1: ingesta de datos"), hd("Sprint 5: pronóstico de precios")],
    tasks,
    ...rows.map((r, k) => [ax(r[0]), cell(r[1], k), cell(r[2], k)]),
  ], { x: 0.5, y: 1.1, w: 9, colW: [1.9, 3.55, 3.55], fontSize: 14, rowH: [0.38, 0.6, 0.46, 0.46, 0.46, 0.46, 0.46], border: { type: "solid", pt: 0.75, color: "C9DCCF" }, valign: "middle", margin: [0.02, 0.08, 0.02, 0.08] });
  s.addShape(pres.shapes.STAR_5_POINT, { x: 8.95, y: 0.4, w: 0.55, h: 0.55, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "Estrella de la retrospectiva" });
  s.addNotes("Estrella de la Retrospectiva aplicada a dos sprints técnicos: el de ingesta de datos y el de pronóstico. Las acciones anticipan las dificultades más probables de cada etapa.");
}

// ---------- 13. Cierre ----------
pres.addSection({ title: "Cierre" });
{
  const s = pres.addSlide({ masterName: "CIERRE", sectionTitle: "Cierre" });
  tx(s, "Muchas gracias", { x: 0.5, y: 1.35, w: 9, h: 1.0, fontSize: 48, bold: true, color: C.background1 });
  tx(s, "¿Preguntas?", { x: 0.5, y: 2.35, w: 9, h: 0.7, fontSize: 30, color: C.accent3 });
  tx(s, [
    { text: "Ing. Nicolás Francisco Moguilner Reh", options: { bold: true, breakLine: true } },
    { text: "Director: Dr. Ing. Facundo Adrián Lucianna" },
  ], { x: 0.5, y: 3.6, w: 9, h: 0.8, fontSize: 18, color: C.background1, paraSpaceAfter: 4 });
  s.addNotes("Cierre y preguntas.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
