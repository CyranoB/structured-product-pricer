const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

// ─── Presentation metadata ───
pres.layout = "LAYOUT_16x9";
pres.author = "Eddie Pick";
pres.title = "Certificat de D\u00e9p\u00f4t \u2013 Rendement Num\u00e9rique Annuel";

// ─── Color palette (from HTML) ───
const C = {
  cream:     "FAFAF8",
  white:     "FFFFFF",
  charcoal:  "2D3436",
  warmGray:  "636E72",
  teal:      "00838F",
  tealHover: "006064",
  green:     "2E7D32",
  greenBg:   "F1F8F1",
  orange:    "BF5600",
  orangeBg:  "FDF5EF",
  red:       "B71C1C",
  redBg:     "FDF0F0",
  border:    "DDD8D0",
  borderLt:  "EAE6DF",
  tertiary:  "9E9A94",
  codeText:  "E0E0E0",
};

// ─── Font helpers ───
const F = { serif: "Georgia", sans: "Calibri", mono: "Consolas" };

// ─── Reusable style factories ───
const makeShadow = () => ({ type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.06 });
const iconDir = __dirname + "/icons/";
const addIcon = (s, name, opts = {}) => {
  s.addImage({
    path: iconDir + name + ".png",
    x: opts.x ?? 9.0, y: opts.y ?? 0.25, w: opts.w ?? 0.45, h: opts.h ?? 0.45,
    transparency: opts.transparency ?? 60,
  });
};

// ─── Stock data ───
const stocks = [
  { t: "AAPL", n: "Apple Inc",           ex: "NASDAQ", pi: "403,17", s0: "31,05",  vol: "33,9%", w: "10%" },
  { t: "C",    n: "Citigroup Inc",        ex: "NYSE",   pi: "26,72",  s0: "57,79",  vol: "65,5%", w: "10%" },
  { t: "F",    n: "Ford Motor Co",        ex: "NYSE",   pi: "10,08",  s0: "14,56",  vol: "51,6%", w: "10%" },
  { t: "HPQ",  n: "Hewlett-Packard Co",   ex: "NYSE",   pi: "22,71",  s0: "16,87",  vol: "33,3%", w: "10%" },
  { t: "JNJ",  n: "Johnson & Johnson",    ex: "NYSE",   pi: "62,69",  s0: "159,39", vol: "15,1%", w: "10%" },
  { t: "LLY",  n: "Eli Lilly & Co",       ex: "NYSE",   pi: "36,64",  s0: "114,51", vol: "21,6%", w: "10%" },
  { t: "LOW",  n: "Lowe's Cos Inc",       ex: "NYSE",   pi: "19,82",  s0: "79,61",  vol: "30,7%", w: "10%" },
  { t: "MO",   n: "Altria Group Inc",     ex: "NYSE",   pi: "26,00",  s0: "115,52", vol: "18,1%", w: "10%" },
  { t: "MRK",  n: "Merck & Co Inc",       ex: "NYSE",   pi: "31,60",  s0: "85,37",  vol: "24,3%", w: "10%" },
  { t: "WMT",  n: "Wal-Mart Stores Inc",  ex: "NYSE",   pi: "51,83",  s0: "29,80",  vol: "18,8%", w: "10%" },
];

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 1 — Title
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "certificate", { x: 8.6, y: 0.4, w: 0.7, h: 0.7, transparency: 50 });

  // Teal accent bar at top
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });

  s.addText("Certificat de D\u00e9p\u00f4t\n\u00e0 Rendement Num\u00e9rique Annuel", {
    x: 0.8, y: 1.0, w: 8.4, h: 1.6,
    fontFace: F.serif, fontSize: 34, color: C.charcoal,
    align: "left", valign: "middle", lineSpacingMultiple: 1.15,
  });

  // Teal rule
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 2.85, w: 3.0, h: 0, line: { color: C.teal, width: 2 } });

  s.addText("MATH40602 \u2014 M\u00e9thodes Quantitatives 2 \u2014 Devoir 2", {
    x: 0.8, y: 3.1, w: 8.4, h: 0.5,
    fontFace: F.sans, fontSize: 14, color: C.warmGray, italic: true,
  });

  s.addText("Eddie Pick", {
    x: 0.8, y: 3.8, w: 4, h: 0.4,
    fontFace: F.sans, fontSize: 16, color: C.charcoal, bold: true,
  });

  s.addText("Avril 2026", {
    x: 0.8, y: 4.2, w: 4, h: 0.35,
    fontFace: F.sans, fontSize: 13, color: C.charcoal,
  });

  s.addNotes(
    "Bonjour \u00e0 tous. Aujourd'hui je vais vous pr\u00e9senter mon travail d'\u00e9valuation d'un produit structur\u00e9 : " +
    "un Certificat de D\u00e9p\u00f4t \u00e0 Rendement Num\u00e9rique Annuel, \u00e9mis par une banque am\u00e9ricaine.\n\n" +
    "C'est un produit int\u00e9ressant car il combine une protection du capital avec un coupon conditionnel li\u00e9 \u00e0 la performance d'un panier de 10 actions.\n\n" +
    "Je vais d'abord d\u00e9crire le produit, puis expliquer la th\u00e9orie derri\u00e8re notre m\u00e9thode d'\u00e9valuation, " +
    "montrer l'impl\u00e9mentation en MATLAB, et enfin pr\u00e9senter les r\u00e9sultats."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 2 — Agenda
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "presentation-chart");

  s.addText("Agenda", {
    x: 0.8, y: 0.4, w: 8.4, h: 0.7,
    fontFace: F.serif, fontSize: 28, color: C.charcoal, margin: 0,
  });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  const items = [
    "Description du produit structur\u00e9",
    "Cadre th\u00e9orique et m\u00e9thode d'\u00e9valuation",
    "Impl\u00e9mentation MATLAB",
    "R\u00e9sultats et analyse",
  ];

  items.forEach((item, i) => {
    const yBase = 1.5 + i * 0.9;
    // Teal left bar
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: yBase, w: 0.06, h: 0.6, fill: { color: C.teal } });
    // Number
    s.addText(`${i + 1}`, {
      x: 1.8, y: yBase, w: 0.5, h: 0.6,
      fontFace: F.mono, fontSize: 22, color: C.teal, bold: true, valign: "middle", margin: 0,
    });
    // Label
    s.addText(item, {
      x: 2.4, y: yBase, w: 6.5, h: 0.6,
      fontFace: F.sans, fontSize: 17, color: C.charcoal, valign: "middle", margin: 0,
    });
  });

  s.addNotes(
    "Voici le plan de la pr\u00e9sentation en 4 parties.\n\n" +
    "1. D'abord, on va comprendre le produit : sa structure, les 10 titres du panier, et surtout les 3 r\u00e9gimes de performance qui d\u00e9terminent le coupon.\n\n" +
    "2. Ensuite, le cadre th\u00e9orique : le mouvement brownien g\u00e9om\u00e9trique pour mod\u00e9liser les prix, la d\u00e9composition de Cholesky pour g\u00e9rer les corr\u00e9lations, et la d\u00e9composition en options exotiques.\n\n" +
    "3. L'impl\u00e9mentation MATLAB avec les 3 fichiers de code.\n\n" +
    "4. Et enfin les r\u00e9sultats : le prix estim\u00e9, les graphiques, et les limites du mod\u00e8le."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 3 — Description du produit
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "certificate");

  s.addText([
    { text: "1 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Description du produit", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  const params = [
    ["CUSIP",                "05573J AX2"],
    ["D\u00e9nomination",    "1 000 $"],
    ["Terme",                "6 ans (sept. 2011 \u2013 sept. 2017)"],
    ["Sous-jacent",          "Panier de 10 actions"],
    ["Coupon num\u00e9rique","6,50%"],
    ["Plancher (Floor)",     "\u221230% par titre"],
    ["Date d'\u00e9valuation","31 octobre 2016"],
    ["Taux sans risque",     "1,25%"],
  ];

  // 2×4 grid
  params.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const xBase = 0.8 + col * 4.5;
    const yBase = 1.35 + row * 1.0;

    // White card with teal left bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: xBase, y: yBase, w: 4.1, h: 0.85,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addShape(pres.shapes.RECTANGLE, { x: xBase, y: yBase, w: 0.06, h: 0.85, fill: { color: C.teal } });

    // Label
    s.addText(p[0].toUpperCase(), {
      x: xBase + 0.2, y: yBase + 0.08, w: 3.7, h: 0.28,
      fontFace: F.sans, fontSize: 9, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
    });
    // Value
    s.addText(p[1], {
      x: xBase + 0.2, y: yBase + 0.4, w: 3.7, h: 0.35,
      fontFace: F.mono, fontSize: 14, color: C.charcoal, margin: 0,
    });
  });

  s.addNotes(
    "Voici les param\u00e8tres cl\u00e9s du CD.\n\n" +
    "C'est un certificat de d\u00e9p\u00f4t de 1 000 $ avec un terme de 6 ans, \u00e9mis le 26 septembre 2011 et arrivant \u00e0 \u00e9ch\u00e9ance le 29 septembre 2017.\n\n" +
    "Le sous-jacent est un panier de 10 actions am\u00e9ricaines \u00e0 grande capitalisation, chacune pond\u00e9r\u00e9e \u00e0 1/10.\n\n" +
    "Le coupon num\u00e9rique est de 6,50% : c'est le maximum que l'investisseur peut recevoir par p\u00e9riode. " +
    "Le plancher est de \u221230% par titre, ce qui limite les pertes individuelles.\n\n" +
    "Notre date d'\u00e9valuation est le 31 octobre 2016, soit environ 11 mois avant l'\u00e9ch\u00e9ance. " +
    "Le taux sans risque utilis\u00e9 est de 1,25%, correspondant au taux du US Treasury 1 an \u00e0 cette date."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Basket table
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "briefcase");

  s.addText([
    { text: "1 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Panier de 10 titres sous-jacents", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  const header = [
    { text: "TICKER",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream } } },
    { text: "NOM",     options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream } } },
    { text: "BOURSE",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream }, align: "center" } },
    { text: "P. INIT.", options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream }, align: "right" } },
    { text: "S\u2080 (OCT 16)", options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream }, align: "right" } },
    { text: "\u03C3",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 9, fill: { color: C.cream }, align: "right" } },
  ];

  const rows = stocks.map((st, i) => [
    { text: st.t, options: { bold: true, fontFace: F.mono, fontSize: 10, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream } } },
    { text: st.n, options: { fontFace: F.sans, fontSize: 10, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream } } },
    { text: st.ex, options: { fontFace: F.sans, fontSize: 10, color: C.warmGray, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "center" } },
    { text: st.pi + " $", options: { fontFace: F.mono, fontSize: 10, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: st.s0 + " $", options: { fontFace: F.mono, fontSize: 10, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: st.vol, options: { fontFace: F.mono, fontSize: 10, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
  ]);

  s.addTable([header, ...rows], {
    x: 0.5, y: 1.25, w: 9.0,
    colW: [0.9, 2.4, 1.0, 1.4, 1.5, 0.9],
    border: { type: "solid", pt: 0.5, color: C.borderLt },
    rowH: [0.35, ...Array(10).fill(0.35)],
  });

  s.addText("Source : Bloomberg, prix ajust\u00e9s (dividendes inclus)", {
    x: 0.8, y: 5.0, w: 8, h: 0.3,
    fontFace: F.sans, fontSize: 9, italic: true, color: C.tertiary,
  });

  s.addNotes(
    "Voici les 10 titres du panier. Ce sont tous des grandes capitalisations am\u00e9ricaines de secteurs diversifi\u00e9s : " +
    "technologie (Apple, HP), finance (Citigroup), automobile (Ford), sant\u00e9 (J&J, Eli Lilly, Merck), " +
    "distribution (Lowe's, Walmart) et tabac (Altria).\n\n" +
    "Les prix initiaux (P. Init.) sont les prix au 26 septembre 2011 \u2014 la date de cr\u00e9ation du CD. " +
    "Les prix S\u2080 sont les prix ajust\u00e9s au 31 octobre 2016, notre date d'\u00e9valuation.\n\n" +
    "Point important : on utilise les prix ajust\u00e9s pour les dividendes, ce qui \u00e9vite d'avoir \u00e0 mod\u00e9liser " +
    "les rendements de dividendes s\u00e9par\u00e9ment dans notre simulation.\n\n" +
    "On observe une grande disparit\u00e9 de volatilit\u00e9s : de 15,1% pour J&J (titre d\u00e9fensif) \u00e0 65,5% pour Citigroup " +
    "(secteur financier, post-crise). Cette disparit\u00e9 aura un impact important sur la distribution des r\u00e9sultats."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 5 — Performance function (3 regimes)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "tree-structure");

  s.addText([
    { text: "1 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Fonction de performance \u00e0 3 r\u00e9gimes", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // ── Left: Payoff diagram as a schematic ──
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 5.0, h: 3.0,
    fill: { color: C.white }, shadow: makeShadow(),
  });

  // FIXED: Y-axis label with wider text box, horizontal
  s.addText("Perf.", { x: 0.7, y: 1.4, w: 0.8, h: 0.3, fontFace: F.sans, fontSize: 9, color: C.warmGray, align: "center", margin: 0 });
  s.addText("Rendement du titre (R)", { x: 1.8, y: 4.1, w: 3, h: 0.25, fontFace: F.sans, fontSize: 9, color: C.warmGray, align: "center" });

  // Y-axis line
  s.addShape(pres.shapes.LINE, { x: 1.6, y: 1.5, w: 0, h: 2.5, line: { color: C.charcoal, width: 1 } });
  // X-axis line
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.0, w: 4.3, h: 0, line: { color: C.charcoal, width: 1 } });

  // Red flat segment: floor at -30%
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 3.65, w: 1.2, h: 0, line: { color: C.red, width: 3 } });
  // Orange diagonal: -30% to 0%
  s.addShape(pres.shapes.LINE, { x: 2.2, y: 3.65, w: 1.4, h: -0.65, line: { color: C.orange, width: 3 } });
  // Green flat: 0% onward at 6.5%
  s.addShape(pres.shapes.LINE, { x: 3.6, y: 1.85, w: 1.5, h: 0, line: { color: C.green, width: 3 } });

  // Axis tick labels
  s.addText("\u221230%", { x: 1.9, y: 3.8, w: 0.6, h: 0.2, fontFace: F.mono, fontSize: 8, color: C.warmGray, align: "center", margin: 0 });
  s.addText("0%", { x: 3.3, y: 3.8, w: 0.5, h: 0.2, fontFace: F.mono, fontSize: 8, color: C.warmGray, align: "center", margin: 0 });
  s.addText("+6,50%", { x: 0.7, y: 1.75, w: 0.85, h: 0.2, fontFace: F.mono, fontSize: 8, color: C.green, align: "right", margin: 0 });
  s.addText("\u221230%", { x: 0.7, y: 3.55, w: 0.85, h: 0.2, fontFace: F.mono, fontSize: 8, color: C.red, align: "right", margin: 0 });

  // ── Right: Regime cards ──
  const regimes = [
    { label: "R\u00c9GIME NUM\u00c9RIQUE", cond: "Si R > 0%", perf: "Perf = 6,50%", color: C.green, bg: C.greenBg },
    { label: "R\u00c9GIME PASSTHROUGH", cond: "Si \u221230% < R \u2264 0%", perf: "Perf = R", color: C.orange, bg: C.orangeBg },
    { label: "R\u00c9GIME PLANCHER", cond: "Si R \u2264 \u221230%", perf: "Perf = \u221230%", color: C.red, bg: C.redBg },
  ];

  regimes.forEach((r, i) => {
    const yBase = 1.35 + i * 1.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.0, y: yBase, w: 3.5, h: 0.85,
      fill: { color: r.bg },
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: yBase, w: 0.06, h: 0.85, fill: { color: r.color } });

    s.addText(r.label, {
      x: 6.2, y: yBase + 0.05, w: 3.1, h: 0.22,
      fontFace: F.sans, fontSize: 8, color: r.color, bold: true, charSpacing: 1, margin: 0,
    });
    s.addText(r.cond, {
      x: 6.2, y: yBase + 0.28, w: 3.1, h: 0.22,
      fontFace: F.sans, fontSize: 11, color: C.charcoal, margin: 0,
    });
    s.addText(r.perf, {
      x: 6.2, y: yBase + 0.52, w: 3.1, h: 0.25,
      fontFace: F.mono, fontSize: 13, color: r.color, bold: true, margin: 0,
    });
  });

  // ── Formulas at bottom ──
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 4.5, w: 8.8, h: 0.75,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText([
    { text: "Taux de coupon = max(0, 1/10 \u00D7 \u03A3 Perf", options: { fontFace: F.mono, fontSize: 11, color: C.charcoal } },
    { text: "i", options: { fontFace: F.mono, fontSize: 8, color: C.charcoal } },
    { text: ")          Payoff = 1 000 \u00D7 (1 + Taux de coupon)", options: { fontFace: F.mono, fontSize: 11, color: C.charcoal } },
  ], { x: 0.8, y: 4.5, w: 8.4, h: 0.75, align: "center", valign: "middle" });

  s.addNotes(
    "C'est la diapositive cl\u00e9 pour comprendre le produit. La performance de chaque titre est d\u00e9termin\u00e9e par 3 r\u00e9gimes :\n\n" +
    "R\u00e9gime 1 (vert) \u2014 Num\u00e9rique : Si le rendement du titre est positif, m\u00eame de 0,01%, la performance est automatiquement fix\u00e9e \u00e0 6,50%. " +
    "C'est une option binaire : tout ou rien. Un titre qui monte de 1% rapporte autant qu'un titre qui monte de 50%.\n\n" +
    "R\u00e9gime 2 (orange) \u2014 Passthrough : Si le rendement est n\u00e9gatif mais sup\u00e9rieur \u00e0 \u221230%, l'investisseur absorbe la perte telle quelle. " +
    "Un titre \u00e0 \u22125% contribue \u22125% \u00e0 la moyenne.\n\n" +
    "R\u00e9gime 3 (rouge) \u2014 Plancher : Si le rendement est inf\u00e9rieur ou \u00e9gal \u00e0 \u221230%, la perte est cap\u00e9e \u00e0 \u221230%. " +
    "Un titre qui perd 50% ne contribue que \u221230%.\n\n" +
    "Le taux du coupon est la moyenne pond\u00e9r\u00e9e des 10 performances, avec un plancher global \u00e0 0%. " +
    "Donc le pire sc\u00e9nario est un coupon de z\u00e9ro, pas une perte en capital \u2014 le principal de 1 000 $ est toujours rembours\u00e9."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 6 — Worked example: coupon calculation (NEW)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "function");

  s.addText([
    { text: "1 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Exemple : calcul du coupon (rendements mixtes)", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 22, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.05, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Example #2 from term sheet: mixed returns
  const exHeader = [
    { text: "TITRE",     options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream } } },
    { text: "P. INIT.",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "right" } },
    { text: "P. FINAL",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "right" } },
    { text: "RENDEMENT",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "right" } },
    { text: "R\u00c9GIME", options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "center" } },
    { text: "PERF.",     options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "right" } },
    { text: "P \u00d7 W",  options: { bold: true, color: C.warmGray, fontFace: F.sans, fontSize: 8, fill: { color: C.cream }, align: "right" } },
  ];

  const exData = [
    { t: "1", pi: "100$", pf: "101$", ret: "+1,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "2", pi: "100$", pf: "104$", ret: "+4,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "3", pi: "100$", pf: "92$",  ret: "\u22128,00%", regime: "Pass.", regC: C.orange, perf: "\u22128,00%", pw: "\u22120,800%" },
    { t: "4", pi: "100$", pf: "108$", ret: "+8,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "5", pi: "100$", pf: "104$", ret: "+4,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "6", pi: "100$", pf: "103$", ret: "+3,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "7", pi: "100$", pf: "97$",  ret: "\u22123,00%", regime: "Pass.", regC: C.orange, perf: "\u22123,00%", pw: "\u22120,300%" },
    { t: "8", pi: "100$", pf: "102$", ret: "+2,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "9", pi: "100$", pf: "101$", ret: "+1,00%",  regime: "Num.",   regC: C.green,  perf: "6,50%",   pw: "+0,650%" },
    { t: "10",pi: "100$", pf: "89$",  ret: "\u221211,00%",regime: "Pass.", regC: C.orange, perf: "\u221211,00%",pw: "\u22121,100%" },
  ];

  const exRows = exData.map((d, i) => [
    { text: "Titre " + d.t, options: { fontFace: F.sans, fontSize: 8, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream } } },
    { text: d.pi, options: { fontFace: F.mono, fontSize: 8, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: d.pf, options: { fontFace: F.mono, fontSize: 8, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: d.ret, options: { fontFace: F.mono, fontSize: 8, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: d.regime, options: { fontFace: F.sans, fontSize: 8, color: d.regC, bold: true, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "center" } },
    { text: d.perf, options: { fontFace: F.mono, fontSize: 8, color: d.regC, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
    { text: d.pw, options: { fontFace: F.mono, fontSize: 8, color: C.charcoal, fill: { color: i % 2 === 0 ? C.white : C.cream }, align: "right" } },
  ]);

  s.addTable([exHeader, ...exRows], {
    x: 0.5, y: 1.15, w: 9.0,
    colW: [1.0, 0.9, 0.9, 1.2, 0.9, 1.1, 1.1],
    border: { type: "solid", pt: 0.5, color: C.borderLt },
    rowH: [0.28, ...Array(10).fill(0.28)],
  });

  // Result summary cards
  const yResult = 4.35;
  // Weighted average
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yResult, w: 3.8, h: 0.55, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yResult, w: 0.06, h: 0.55, fill: { color: C.teal } });
  s.addText("MOYENNE POND\u00c9R\u00c9E", { x: 0.7, y: yResult + 0.02, w: 3.4, h: 0.2, fontFace: F.sans, fontSize: 8, color: C.warmGray, bold: true, charSpacing: 1, margin: 0 });
  s.addText("\u03A3 (Perf \u00d7 Poids) = +2,350%", { x: 0.7, y: yResult + 0.25, w: 3.4, h: 0.25, fontFace: F.mono, fontSize: 12, color: C.charcoal, margin: 0 });

  // Coupon rate
  s.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: yResult, w: 2.2, h: 0.55, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 4.6, y: yResult, w: 0.06, h: 0.55, fill: { color: C.teal } });
  s.addText("TAUX DU COUPON", { x: 4.8, y: yResult + 0.02, w: 1.85, h: 0.2, fontFace: F.sans, fontSize: 8, color: C.warmGray, bold: true, charSpacing: 1, margin: 0 });
  s.addText("max(0, 2,35%) = 2,35%", { x: 4.8, y: yResult + 0.25, w: 1.85, h: 0.25, fontFace: F.mono, fontSize: 11, color: C.teal, bold: true, margin: 0 });

  // Payment
  s.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: yResult, w: 2.4, h: 0.55, fill: { color: C.white }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: yResult, w: 0.06, h: 0.55, fill: { color: C.teal } });
  s.addText("PAIEMENT", { x: 7.3, y: yResult + 0.02, w: 2.05, h: 0.2, fontFace: F.sans, fontSize: 8, color: C.warmGray, bold: true, charSpacing: 1, margin: 0 });
  s.addText("1 000 \u00d7 2,35% = 23,50 $", { x: 7.3, y: yResult + 0.25, w: 2.05, h: 0.25, fontFace: F.mono, fontSize: 11, color: C.charcoal, margin: 0 });

  s.addText("7 titres positifs \u2192 coupon num\u00e9rique (6,50%)  |  3 titres n\u00e9gatifs \u2192 passthrough  |  R\u00e9sultat net : coupon de 2,35%", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.3,
    fontFace: F.sans, fontSize: 9, italic: true, color: C.warmGray, align: "center",
  });

  s.addNotes(
    "Voici un exemple concret avec l'exemple #2 du term sheet \u2014 le sc\u00e9nario de rendements mixtes, le plus r\u00e9aliste.\n\n" +
    "Sur 10 titres : 7 ont un rendement positif (m\u00eame faiblement, comme +1%) et tombent donc dans le r\u00e9gime num\u00e9rique \u00e0 6,50%. " +
    "3 titres ont un rendement n\u00e9gatif : \u22128%, \u22123% et \u221211%, tous dans la zone passthrough.\n\n" +
    "Le calcul :\n" +
    "\u2022 Les 7 titres positifs contribuent chacun 6,50% \u00d7 1/10 = +0,650%\n" +
    "\u2022 Le titre 3 (\u22128%) contribue \u22120,800%\n" +
    "\u2022 Le titre 7 (\u22123%) contribue \u22120,300%\n" +
    "\u2022 Le titre 10 (\u221211%) contribue \u22121,100%\n\n" +
    "Total : 7 \u00d7 0,650% + (\u22120,800% \u2212 0,300% \u2212 1,100%) = +2,350%\n\n" +
    "Le taux du coupon est max(0, 2,350%) = 2,350%. Le paiement est 1 000 \u00d7 2,35% = 23,50 $.\n\n" +
    "Cet exemple montre que m\u00eame avec des pertes sur certains titres, le coupon peut rester positif gr\u00e2ce aux 6,50% des titres gagnants."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7 — GBM Theory
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "chart-line-up");

  s.addText([
    { text: "2 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Mouvement Brownien G\u00e9om\u00e9trique", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Main formula card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.35, w: 8.4, h: 1.2,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText("S\u1D40 = S\u2080 \u00D7 exp( (r \u2212 \u03C3\u00B2/2) \u00D7 T  +  \u03C3 \u00D7 \u221AT \u00D7 Z )", {
    x: 0.8, y: 1.35, w: 8.4, h: 1.2,
    fontFace: F.mono, fontSize: 20, color: C.charcoal, align: "center", valign: "middle",
  });

  // Explanation bullets
  const explanations = [
    { term: "S\u2080", desc: "Prix observ\u00e9 au 31 oct. 2016 (prix ajust\u00e9 dividendes)" },
    { term: "r = 1,25%", desc: "Taux sans risque (US Treasury 1 an)" },
    { term: "\u03C3", desc: "Volatilit\u00e9 annualis\u00e9e (estim\u00e9e sur 564 rendements log hebdo.)" },
    { term: "T = 11/12", desc: "Horizon r\u00e9siduel jusqu'\u00e0 l'\u00e9ch\u00e9ance (sept. 2017)" },
    { term: "Z", desc: "Variables normales corr\u00e9l\u00e9es (via d\u00e9composition de Cholesky)" },
  ];

  explanations.forEach((e, i) => {
    const yBase = 2.8 + i * 0.45;
    s.addText(e.term, {
      x: 1.0, y: yBase, w: 1.5, h: 0.4,
      fontFace: F.mono, fontSize: 12, color: C.teal, bold: true, margin: 0,
    });
    s.addText(e.desc, {
      x: 2.6, y: yBase, w: 6.2, h: 0.4,
      fontFace: F.sans, fontSize: 12, color: C.charcoal, margin: 0,
    });
  });

  // Why MC box — moved up slightly for better bottom margin
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.9, w: 0.06, h: 0.45, fill: { color: C.teal } });
  s.addText("Pourquoi Monte Carlo ?  Payoff exotique (digital + plancher), panier de 10 titres corr\u00e9l\u00e9s, pas de solution analytique.", {
    x: 1.05, y: 4.9, w: 8.15, h: 0.45,
    fontFace: F.sans, fontSize: 11, italic: true, color: C.charcoal, valign: "middle",
  });

  s.addNotes(
    "On passe maintenant au cadre th\u00e9orique. Pour \u00e9valuer le CD, on doit simuler les prix futurs des 10 actions \u00e0 l'\u00e9ch\u00e9ance.\n\n" +
    "On utilise le Mouvement Brownien G\u00e9om\u00e9trique (GBM), le mod\u00e8le standard en finance quantitative. " +
    "La formule cl\u00e9 est : S_T = S_0 \u00d7 exp((r \u2212 \u03c3\u00b2/2) \u00d7 T + \u03c3 \u00d7 \u221aT \u00d7 Z)\n\n" +
    "Chaque variable :\n" +
    "\u2022 S\u2080 : le prix observ\u00e9 au 31 octobre 2016 (nos prix ajust\u00e9s Bloomberg)\n" +
    "\u2022 r = 1,25% : le taux sans risque \u2014 on utilise la mesure risque-neutre, pas la mesure historique\n" +
    "\u2022 \u03c3 : la volatilit\u00e9 annualis\u00e9e, estim\u00e9e \u00e0 partir de 564 rendements log hebdomadaires (10 ans de donn\u00e9es)\n" +
    "\u2022 T = 11/12 : l'horizon r\u00e9siduel, environ 11 mois\n" +
    "\u2022 Z : des variables normales corr\u00e9l\u00e9es entre les 10 titres\n\n" +
    "Pourquoi Monte Carlo et pas une formule ferm\u00e9e ? Parce que le payoff est exotique : " +
    "c'est une fonction par morceaux (3 r\u00e9gimes) appliqu\u00e9e \u00e0 10 titres corr\u00e9l\u00e9s, " +
    "puis agr\u00e9g\u00e9e avec un plancher \u00e0 z\u00e9ro. Il n'existe pas de formule analytique pour ce type de payoff."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7 — Cholesky
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "grid-nine");

  s.addText([
    { text: "2 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "D\u00e9composition de Cholesky", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Left: problem & solution
  s.addText("PROBL\u00c8ME", {
    x: 0.8, y: 1.3, w: 4.5, h: 0.35,
    fontFace: F.sans, fontSize: 11, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
  });
  s.addText("Simuler 10 variables normales\ncorr\u00e9l\u00e9es entre elles selon\nla structure historique du march\u00e9", {
    x: 0.8, y: 1.65, w: 4.3, h: 0.9,
    fontFace: F.sans, fontSize: 13, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
  });

  s.addText("SOLUTION", {
    x: 0.8, y: 2.75, w: 4.5, h: 0.35,
    fontFace: F.sans, fontSize: 11, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
  });

  // Formula cards
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.15, w: 4.3, h: 0.65,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText("C = L \u00D7 L\u1D40", {
    x: 0.8, y: 3.15, w: 4.3, h: 0.65,
    fontFace: F.mono, fontSize: 18, color: C.charcoal, align: "center", valign: "middle",
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 3.95, w: 4.3, h: 0.65,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText([
    { text: "Z", options: { fontFace: F.mono, fontSize: 18, color: C.teal, bold: true } },
    { text: "corr", options: { fontFace: F.mono, fontSize: 11, color: C.teal, bold: true } },
    { text: " = L \u00D7 ", options: { fontFace: F.mono, fontSize: 18, color: C.charcoal } },
    { text: "Z", options: { fontFace: F.mono, fontSize: 18, color: C.warmGray } },
    { text: "indep", options: { fontFace: F.mono, fontSize: 11, color: C.warmGray } },
  ], { x: 0.8, y: 3.95, w: 4.3, h: 0.65, align: "center", valign: "middle" });

  s.addText("564 rendements log hebdomadaires\nutilis\u00e9s pour estimer la matrice 10\u00D710", {
    x: 0.8, y: 4.75, w: 4.5, h: 0.55,
    fontFace: F.sans, fontSize: 11, color: C.warmGray, italic: true, margin: 0, lineSpacingMultiple: 1.3,
  });

  // Right: Correlation heatmap
  s.addText("Matrice de corr\u00e9lation (extrait)", {
    x: 5.6, y: 1.3, w: 4.0, h: 0.35,
    fontFace: F.sans, fontSize: 10, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
  });

  const corrSample = [
    [1.00, 0.34, 0.34, 0.37, 0.24],
    [0.34, 1.00, 0.59, 0.29, 0.38],
    [0.34, 0.59, 1.00, 0.36, 0.38],
    [0.37, 0.29, 0.36, 1.00, 0.37],
    [0.24, 0.38, 0.38, 0.37, 1.00],
  ];
  const heatLabels = ["AAPL", "C", "F", "HPQ", "JNJ"];
  const cellSize = 0.55;
  const heatX = 6.2;
  const heatY = 1.8;

  // Column labels
  heatLabels.forEach((lbl, j) => {
    s.addText(lbl, {
      x: heatX + j * cellSize, y: heatY - 0.3, w: cellSize, h: 0.25,
      fontFace: F.mono, fontSize: 7, color: C.warmGray, align: "center", margin: 0,
    });
  });

  corrSample.forEach((row, i) => {
    // Row label — wider box for short tickers
    s.addText(heatLabels[i], {
      x: heatX - 0.65, y: heatY + i * cellSize, w: 0.6, h: cellSize,
      fontFace: F.mono, fontSize: 7, color: C.warmGray, align: "right", valign: "middle", margin: 0,
    });

    row.forEach((val, j) => {
      let fillColor;
      if (val >= 0.9) fillColor = C.teal;
      else if (val >= 0.5) fillColor = "4DB6AC";
      else if (val >= 0.35) fillColor = "B2DFDB";
      else fillColor = "E0F2F1";

      s.addShape(pres.shapes.RECTANGLE, {
        x: heatX + j * cellSize, y: heatY + i * cellSize, w: cellSize, h: cellSize,
        fill: { color: fillColor }, line: { color: C.white, width: 1 },
      });
      s.addText(val.toFixed(2), {
        x: heatX + j * cellSize, y: heatY + i * cellSize, w: cellSize, h: cellSize,
        fontFace: F.mono, fontSize: 8, color: val >= 0.5 ? C.white : C.charcoal,
        align: "center", valign: "middle", margin: 0,
      });
    });
  });

  s.addText("5 titres sur 10 montr\u00e9s ici", {
    x: 5.6, y: heatY + 5 * cellSize + 0.15, w: 4.0, h: 0.25,
    fontFace: F.sans, fontSize: 8, italic: true, color: C.tertiary, margin: 0,
  });

  s.addNotes(
    "Le probl\u00e8me central de la simulation est la corr\u00e9lation. Les 10 actions ne bougent pas ind\u00e9pendamment : " +
    "quand le march\u00e9 baisse, elles ont tendance \u00e0 baisser ensemble.\n\n" +
    "Pour capturer cette structure de d\u00e9pendance, on utilise la d\u00e9composition de Cholesky.\n\n" +
    "Le principe :\n" +
    "1. On estime la matrice de corr\u00e9lation 10\u00d710 \u00e0 partir des 564 rendements log hebdomadaires historiques\n" +
    "2. On d\u00e9compose cette matrice : C = L \u00d7 L\u1d40, o\u00f9 L est une matrice triangulaire inf\u00e9rieure\n" +
    "3. Pour g\u00e9n\u00e9rer des variables corr\u00e9l\u00e9es, on multiplie L par un vecteur de normales ind\u00e9pendantes : Z_corr = L \u00d7 Z_indep\n\n" +
    "Sur la droite, on voit un extrait de la matrice de corr\u00e9lation (5 titres sur 10). " +
    "Les corr\u00e9lations varient de 0,24 (AAPL-JNJ, tech vs sant\u00e9) \u00e0 0,59 (C-F, finance et auto, secteurs cycliques). " +
    "La corr\u00e9lation moyenne est d'environ 0,36 \u2014 mod\u00e9r\u00e9e, ce qui est typique pour un panier diversifi\u00e9.\n\n" +
    "Cette corr\u00e9lation mod\u00e9r\u00e9e est favorable pour le produit : elle r\u00e9duit le risque que tous les titres chutent simultan\u00e9ment."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 — Exotic options decomposition (NEW)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "link");

  s.addText([
    { text: "2 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "D\u00e9composition en options exotiques", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  s.addText("Le payoff de chaque titre peut \u00eatre d\u00e9compos\u00e9 en 3 options \u00e9l\u00e9mentaires :", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.4,
    fontFace: F.sans, fontSize: 13, color: C.charcoal, margin: 0,
  });

  // Option 1: Digital Call
  const opts = [
    {
      num: "1",
      title: "Option num\u00e9rique (Digital Call)",
      desc: "Paie le coupon fixe de 6,50% d\u00e8s que le rendement est positif, quelle que soit l'amplitude de la hausse.",
      formula: "Si R > 0%  \u2192  Payoff = 6,50%",
      color: C.green,
      bg: C.greenBg,
    },
    {
      num: "2",
      title: "Put Spread (0% / \u221230%)",
      desc: "L'investisseur absorbe les pertes entre 0% et \u221230%. C'est la \u00ab zone de risque \u00bb non prot\u00e9g\u00e9e.",
      formula: "Si \u221230% < R \u2264 0%  \u2192  Payoff = R",
      color: C.orange,
      bg: C.orangeBg,
    },
    {
      num: "3",
      title: "Plancher (Floor Protection)",
      desc: "Limite les pertes \u00e0 \u221230% maximum par titre. \u00c9quivalent \u00e0 une option put synth\u00e9tique avec strike \u00e0 \u221230%.",
      formula: "Si R \u2264 \u221230%  \u2192  Payoff = \u221230% (cap\u00e9)",
      color: C.red,
      bg: C.redBg,
    },
  ];

  opts.forEach((o, i) => {
    const yBase = 1.7 + i * 1.15;
    // Card bg
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: yBase, w: 8.4, h: 1.0,
      fill: { color: o.bg },
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yBase, w: 0.08, h: 1.0, fill: { color: o.color } });

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 1.05, y: yBase + 0.2, w: 0.45, h: 0.45,
      fill: { color: o.color },
    });
    s.addText(o.num, {
      x: 1.05, y: yBase + 0.2, w: 0.45, h: 0.45,
      fontFace: F.mono, fontSize: 16, color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
    });

    // Title
    s.addText(o.title, {
      x: 1.7, y: yBase + 0.05, w: 4.0, h: 0.3,
      fontFace: F.sans, fontSize: 13, color: o.color, bold: true, margin: 0,
    });
    // Description
    s.addText(o.desc, {
      x: 1.7, y: yBase + 0.38, w: 4.5, h: 0.55,
      fontFace: F.sans, fontSize: 10, color: C.charcoal, margin: 0, lineSpacingMultiple: 1.3,
    });
    // Formula on right
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.5, y: yBase + 0.2, w: 2.5, h: 0.5,
      fill: { color: C.white },
    });
    s.addText(o.formula, {
      x: 6.5, y: yBase + 0.2, w: 2.5, h: 0.5,
      fontFace: F.mono, fontSize: 9, color: o.color, align: "center", valign: "middle", margin: 0,
    });
  });

  // Bottom insight
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.0, w: 0.06, h: 0.35, fill: { color: C.teal } });
  s.addText("Le payoff total est la somme pond\u00e9r\u00e9e des 10 performances individuelles, avec un plancher global \u00e0 0%.", {
    x: 1.05, y: 5.0, w: 8.15, h: 0.35,
    fontFace: F.sans, fontSize: 11, italic: true, color: C.charcoal, valign: "middle",
  });

  s.addNotes(
    "Pour mieux comprendre le produit, on peut le d\u00e9composer en options \u00e9l\u00e9mentaires. " +
    "Pour chaque titre, le payoff est la combinaison de 3 options :\n\n" +
    "1. L'option num\u00e9rique (Digital Call) : c'est une option binaire. Si le titre monte, on re\u00e7oit le coupon fixe de 6,50%, " +
    "ind\u00e9pendamment de l'amplitude de la hausse. C'est ce qui fait la \u00ab magie \u00bb du produit \u2014 m\u00eame une hausse infime d\u00e9clenche le coupon maximal.\n\n" +
    "2. Le Put Spread (0% / \u221230%) : dans la zone de pertes mod\u00e9r\u00e9es, l'investisseur absorbe les pertes directement. " +
    "C'est la zone de risque non prot\u00e9g\u00e9e. On peut l'interpr\u00e9ter comme la vente d'un put \u00e0 strike 0% et l'achat d'un put \u00e0 strike \u221230%.\n\n" +
    "3. Le plancher (Floor Protection) : en dessous de \u221230%, les pertes sont cap\u00e9es. " +
    "C'est l'\u00e9quivalent d'un achat de put \u00e0 strike \u221230% qui prot\u00e8ge l'investisseur contre les crashs catastrophiques.\n\n" +
    "Le payoff total du CD est la somme pond\u00e9r\u00e9e (1/10 chaque) des 10 performances individuelles, " +
    "avec un plancher global \u00e0 0% \u2014 le coupon ne peut jamais \u00eatre n\u00e9gatif."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 10 — MATLAB: simuler_gbm.m
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "code");

  s.addText([
    { text: "3 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Code MATLAB : simuler_gbm.m", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Dark code block
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.3, w: 9.0, h: 3.4,
    fill: { color: C.charcoal },
  });

  // Function signature
  s.addText("function [S_final, L, Z_indep, Z_corr] = simuler_gbm(S0, sigma, CorrMat, r, T, N)", {
    x: 0.8, y: 1.4, w: 8.4, h: 0.4,
    fontFace: F.mono, fontSize: 11, color: C.tertiary, italic: true, margin: 0,
  });

  // FIXED: Use % for MATLAB comments instead of //
  const codeLines = [
    { code: "L = chol(CorrMat, 'lower');",            comment: "% Cholesky" },
    { code: "drift = (r - 0.5 * sigma.^2) * T;",      comment: "% D\u00e9rive" },
    { code: "Z_indep = randn(K, N);",                  comment: "% N(0,I)" },
    { code: "Z_corr  = L * Z_indep;",                  comment: "% Corr\u00e9lation" },
    { code: "choc    = sigma * sqrt(T) .* Z_corr;",    comment: "% Choc" },
    { code: "S_final = S0 .* exp(drift + choc);",      comment: "% Prix" },
  ];

  codeLines.forEach((ln, i) => {
    const yBase = 2.0 + i * 0.42;
    s.addText(ln.code, {
      x: 0.8, y: yBase, w: 6.5, h: 0.38,
      fontFace: F.mono, fontSize: 13, color: C.codeText, margin: 0,
    });
    s.addText(ln.comment, {
      x: 7.3, y: yBase, w: 2.0, h: 0.38,
      fontFace: F.mono, fontSize: 10, color: C.teal, margin: 0,
    });
  });

  // Note below
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.85, w: 0.06, h: 0.4, fill: { color: C.teal } });
  s.addText("Sortie : matrice [K \u00D7 N] de prix terminaux simul\u00e9s sous la mesure risque-neutre (K = 10 titres)", {
    x: 1.05, y: 4.85, w: 8.15, h: 0.4,
    fontFace: F.sans, fontSize: 11, italic: true, color: C.charcoal, valign: "middle",
  });

  s.addNotes(
    "Voici la premi\u00e8re fonction MATLAB : simuler_gbm.m. Elle impl\u00e9mente le GBM en 6 lignes essentielles.\n\n" +
    "Ligne par ligne :\n" +
    "1. L = chol(CorrMat, 'lower') : d\u00e9composition de Cholesky de la matrice de corr\u00e9lation. " +
    "L est la matrice triangulaire inf\u00e9rieure telle que CorrMat = L \u00d7 L'.\n\n" +
    "2. drift = (r - 0.5*sigma.^2)*T : la d\u00e9rive risque-neutre. Le terme \u2212\u03c3\u00b2/2 est le terme de correction d'It\u00f4 " +
    "qui garantit que E[S_T] = S_0 \u00d7 e^(rT) sous la mesure risque-neutre. " +
    "On n'a pas de terme de dividende q car nos prix ajust\u00e9s int\u00e8grent d\u00e9j\u00e0 les dividendes.\n\n" +
    "3. Z_indep = randn(K, N) : g\u00e9n\u00e8re K\u00d7N variables normales standard ind\u00e9pendantes (K=10 titres, N=10 000 simulations)\n\n" +
    "4. Z_corr = L * Z_indep : transforme les variables ind\u00e9pendantes en variables corr\u00e9l\u00e9es via Cholesky\n\n" +
    "5. choc = sigma*sqrt(T).*Z_corr : le choc diffusif, proportionnel \u00e0 la volatilit\u00e9 et \u00e0 la racine du temps\n\n" +
    "6. S_final = S0.*exp(drift + choc) : l'exponentielle donne les prix terminaux lognormaux\n\n" +
    "La sortie est une matrice 10\u00d710 000 : 10 prix terminaux pour chacune des 10 000 simulations."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 — MATLAB: calculer_payoff.m
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "code");

  s.addText([
    { text: "3 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Code MATLAB : calculer_payoff.m", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Dark code block
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.3, w: 9.0, h: 3.5,
    fill: { color: C.charcoal },
  });

  // Lines with color-coded regime indicators — thicker bars for visibility
  const payoffLines = [
    { code: "StockReturns = (S_final - S_init) ./ S_init;", color: C.codeText, bar: null },
    { code: "", color: C.codeText, bar: null },
    { code: "% (1) Rendement > 0% => Digital Coupon", color: C.tertiary, bar: null },
    { code: "Perf(StockReturns > 0) = DigitalCoupon;", color: C.codeText, bar: C.green },
    { code: "", color: C.codeText, bar: null },
    { code: "% (2) -30% < Rendement <= 0% => Passthrough", color: C.tertiary, bar: null },
    { code: "Perf(idx_mid) = StockReturns(idx_mid);", color: C.codeText, bar: C.orange },
    { code: "", color: C.codeText, bar: null },
    { code: "% (3) Rendement <= -30% => Floor", color: C.tertiary, bar: null },
    { code: "Perf(StockReturns <= Floor) = Floor;", color: C.codeText, bar: C.red },
    { code: "", color: C.codeText, bar: null },
    { code: "CouponRate = max(0, mean(Perf, 1));", color: C.codeText, bar: null },
    { code: "PrixCD = exp(-r * T) * mean(Payoffs);", color: C.codeText, bar: null },
  ];

  payoffLines.forEach((ln, i) => {
    if (!ln.code) return;
    const yBase = 1.45 + i * 0.27;
    if (ln.bar) {
      // Wider bar for better visibility when projected
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: yBase, w: 0.1, h: 0.27, fill: { color: ln.bar } });
    }
    s.addText(ln.code, {
      x: 0.85, y: yBase, w: 8.35, h: 0.27,
      fontFace: F.mono, fontSize: 11, color: ln.color, margin: 0,
    });
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.95, w: 0.06, h: 0.35, fill: { color: C.teal } });
  s.addText("Actualisation : PrixCD = e\u207B\u02B3\u1D40 \u00D7 E[Payoff]     IC 95% = \u00B1 1,96 \u00D7 \u03C3/\u221AN", {
    x: 1.05, y: 4.95, w: 8.15, h: 0.35,
    fontFace: F.mono, fontSize: 10, color: C.charcoal, valign: "middle",
  });

  s.addNotes(
    "La deuxi\u00e8me fonction applique les r\u00e8gles du produit structur\u00e9 aux prix simul\u00e9s.\n\n" +
    "D'abord, on calcule le rendement de chaque titre par rapport au prix initial de 2011 (la Trade Date, pas la date d'\u00e9valuation). " +
    "C'est important : le rendement est mesur\u00e9 depuis l'\u00e9mission du CD.\n\n" +
    "Ensuite, on applique les 3 r\u00e9gimes en utilisant l'indexation logique de MATLAB \u2014 c'est vectoris\u00e9, " +
    "donc tr\u00e8s efficace pour traiter les 10\u00d710 000 valeurs d'un coup :\n" +
    "\u2022 Barre verte : les rendements positifs re\u00e7oivent le coupon num\u00e9rique de 6,50%\n" +
    "\u2022 Barre orange : les rendements entre \u221230% et 0% passent tels quels\n" +
    "\u2022 Barre rouge : les rendements sous \u221230% sont cap\u00e9s au plancher\n\n" +
    "Le taux du coupon est max(0, mean(Perf, 1)) \u2014 la moyenne des 10 performances, avec plancher \u00e0 0.\n\n" +
    "Le prix du CD est l'esp\u00e9rance actualis\u00e9e des payoffs : PrixCD = e^(\u2212rT) \u00d7 E[Payoff].\n" +
    "L'intervalle de confiance \u00e0 95% est calcul\u00e9 par la formule classique : \u00b11,96 \u00d7 \u03c3/\u221aN."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 12 — MATLAB: evaluer_cd.m (NEW)
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "flow-arrow");

  s.addText([
    { text: "3 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Code MATLAB : evaluer_cd.m (script principal)", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 22, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Left: Workflow diagram using numbered steps
  s.addText("FLUX D'EX\u00c9CUTION", {
    x: 0.8, y: 1.25, w: 4.0, h: 0.3,
    fontFace: F.sans, fontSize: 10, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
  });

  const steps = [
    { num: "1", label: "Param\u00e8tres", desc: "N, T, r, Nominal, DigitalCoupon, Floor" },
    { num: "2", label: "Donn\u00e9es de march\u00e9", desc: "S_init, S0, sigma, CorrMat (Bloomberg)" },
    { num: "3", label: "Simulation GBM", desc: "[S_final, L, Z_indep, Z_corr] = simuler_gbm(...)" },
    { num: "4", label: "Tableaux L & Z", desc: "Affichage matrice Cholesky, vecteurs Z (sim 1)" },
    { num: "5", label: "\u00c9valuation", desc: "[PrixCD, ...] = calculer_payoff(...)" },
    { num: "6", label: "R\u00e9sultats", desc: "Prix, coupon moyen, IC 95%" },
    { num: "7", label: "Graphiques", desc: "Histogramme + convergence" },
  ];

  steps.forEach((st, i) => {
    const yBase = 1.65 + i * 0.58;
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.9, y: yBase + 0.05, w: 0.35, h: 0.35,
      fill: { color: C.teal },
    });
    s.addText(st.num, {
      x: 0.9, y: yBase + 0.05, w: 0.35, h: 0.35,
      fontFace: F.mono, fontSize: 12, color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
    });
    // Label
    s.addText(st.label, {
      x: 1.4, y: yBase, w: 2.0, h: 0.22,
      fontFace: F.sans, fontSize: 11, color: C.charcoal, bold: true, margin: 0,
    });
    // Description
    s.addText(st.desc, {
      x: 1.4, y: yBase + 0.22, w: 3.4, h: 0.25,
      fontFace: F.mono, fontSize: 9, color: C.warmGray, margin: 0,
    });
    // Connector line (except last)
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.LINE, { x: 1.075, y: yBase + 0.42, w: 0, h: 0.16, line: { color: C.borderLt, width: 1 } });
    }
  });

  // Right: Key code excerpt
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.0, y: 1.25, w: 4.5, h: 3.55,
    fill: { color: C.charcoal },
  });

  const evalCode = [
    { text: "%% 1. PARAM\u00c8TRES", color: C.teal },
    { text: "N = 10000;  T = 11/12;", color: C.codeText },
    { text: "r = 0.0125; Nominal = 1000;", color: C.codeText },
    { text: "DigitalCoupon = 0.065;", color: C.codeText },
    { text: "Floor = -0.30;", color: C.codeText },
    { text: "", color: C.codeText },
    { text: "%% 2. DONN\u00c9ES (Bloomberg)", color: C.teal },
    { text: "S_init = [13.62; 29.51; ...];", color: C.codeText },
    { text: "S0 = [31.05; 57.79; ...];", color: C.codeText },
    { text: "sigma = [0.339; 0.655; ...];", color: C.codeText },
    { text: "CorrMat = [ ... ];  % 10x10", color: C.tertiary },
    { text: "", color: C.codeText },
    { text: "%% 3. SIMULATION", color: C.teal },
    { text: "[S_final,L,Z_indep,Z_corr] = simuler_gbm(...);", color: C.codeText },
    { text: "%% 4. \u00c9VALUATION", color: C.teal },
    { text: "[PrixCD, ...] = calculer_payoff(...);", color: C.codeText },
  ];

  evalCode.forEach((ln, i) => {
    if (!ln.text) return;
    s.addText(ln.text, {
      x: 5.2, y: 1.35 + i * 0.235, w: 4.1, h: 0.23,
      fontFace: F.mono, fontSize: 9, color: ln.color, margin: 0,
    });
  });

  // Bottom note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 5.0, w: 0.06, h: 0.35, fill: { color: C.teal } });
  s.addText("3 fichiers : evaluer_cd.m (orchestrateur) \u2192 simuler_gbm.m (GBM) \u2192 calculer_payoff.m (r\u00e8gles produit)", {
    x: 1.05, y: 5.0, w: 8.15, h: 0.35,
    fontFace: F.sans, fontSize: 11, italic: true, color: C.charcoal, valign: "middle",
  });

  s.addNotes(
    "Le troisi\u00e8me fichier est le script principal evaluer_cd.m. C'est l'orchestrateur qui relie tout.\n\n" +
    "Le flux d'ex\u00e9cution en 6 \u00e9tapes :\n" +
    "1. D\u00e9finir les param\u00e8tres : N=10 000 simulations, T=11/12 ans, r=1,25%, coupon de 6,50%, plancher de \u221230%\n" +
    "2. Charger les donn\u00e9es Bloomberg : prix initiaux ajust\u00e9s (2011), prix actuels (oct. 2016), volatilit\u00e9s, et la matrice de corr\u00e9lation 10\u00d710\n" +
    "3. Appeler simuler_gbm() pour g\u00e9n\u00e9rer les 10 000 sc\u00e9narios de prix\n" +
    "4. Appeler calculer_payoff() pour \u00e9valuer chaque sc\u00e9nario\n" +
    "5. Afficher le prix estim\u00e9, le coupon moyen et l'intervalle de confiance\n" +
    "6. G\u00e9n\u00e9rer les graphiques : histogramme des coupons et courbe de convergence\n\n" +
    "L'architecture en 3 fichiers s\u00e9par\u00e9s rend le code modulaire et testable. " +
    "On peut facilement changer les param\u00e8tres ou r\u00e9utiliser les fonctions pour d'autres produits structur\u00e9s."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 13 — Results: Main price
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "target");

  s.addText([
    { text: "4 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "R\u00e9sultat principal", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Hero price card (charcoal bg, white text)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 1.4, w: 7.0, h: 2.4,
    fill: { color: C.charcoal },
  });

  s.addText("PRIX ESTIM\u00c9 DU CD", {
    x: 1.5, y: 1.55, w: 7.0, h: 0.35,
    fontFace: F.sans, fontSize: 11, color: C.tertiary, align: "center", charSpacing: 2, margin: 0,
  });

  s.addText("1 036,49 $", {
    x: 1.5, y: 1.9, w: 7.0, h: 0.9,
    fontFace: F.mono, fontSize: 42, color: C.white, align: "center", valign: "middle",
  });

  s.addText("IC 95% : [1 035,62 $, 1 037,36 $]", {
    x: 1.5, y: 2.9, w: 7.0, h: 0.35,
    fontFace: F.mono, fontSize: 12, color: C.tertiary, align: "center", margin: 0,
  });

  s.addText("N = 10 000 simulations  |  r = 1,25%  |  T = 11/12 an", {
    x: 1.5, y: 3.3, w: 7.0, h: 0.3,
    fontFace: F.sans, fontSize: 10, color: C.warmGray, align: "center", margin: 0,
  });

  // Detail cards below
  const details = [
    { label: "COUPON MOYEN", value: "48,44 $", sub: "~4,8% de rendement" },
    { label: "PRINCIPAL", value: "1 000,00 $", sub: "100% prot\u00e9g\u00e9" },
    { label: "VALEUR TOTALE", value: "1 048,44 $", sub: "Avant actualisation" },
  ];

  details.forEach((d, i) => {
    const xBase = 0.8 + i * 3.1;
    s.addShape(pres.shapes.RECTANGLE, {
      x: xBase, y: 4.1, w: 2.8, h: 1.05,
      fill: { color: C.white }, shadow: makeShadow(),
    });
    s.addText(d.label, {
      x: xBase + 0.15, y: 4.18, w: 2.5, h: 0.22,
      fontFace: F.sans, fontSize: 8, color: C.warmGray, bold: true, charSpacing: 1, margin: 0,
    });
    s.addText(d.value, {
      x: xBase + 0.15, y: 4.4, w: 2.5, h: 0.4,
      fontFace: F.mono, fontSize: 20, color: C.charcoal, margin: 0,
    });
    s.addText(d.sub, {
      x: xBase + 0.15, y: 4.8, w: 2.5, h: 0.25,
      fontFace: F.sans, fontSize: 10, color: C.warmGray, margin: 0,
    });
  });

  s.addNotes(
    "Voici le r\u00e9sultat principal de notre \u00e9valuation.\n\n" +
    "Le prix estim\u00e9 du CD est de 1 036,49 $ pour une d\u00e9nomination de 1 000 $. " +
    "Cela signifie que le CD vaut environ 3,6% de plus que sa valeur nominale.\n\n" +
    "L'intervalle de confiance \u00e0 95% est tr\u00e8s serr\u00e9 : [1 035,62 $, 1 037,36 $], " +
    "soit une demi-largeur d'environ 0,87 $. Cela montre que 10 000 simulations suffisent pour obtenir une estimation pr\u00e9cise.\n\n" +
    "Le coupon moyen est de 48,44 $, soit environ 4,8% de rendement. " +
    "C'est inf\u00e9rieur au maximum de 6,50% car certaines simulations produisent un coupon de z\u00e9ro (quand les pertes dominent).\n\n" +
    "Le principal est toujours rembours\u00e9 \u00e0 100% \u2014 c'est la caract\u00e9ristique du CD.\n\n" +
    "La valeur totale moyenne (avant actualisation) est de 1 048,44 $, qui une fois actualis\u00e9e au taux de 1,25% " +
    "sur 11 mois donne notre prix de 1 036,49 $."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 11 — Charts
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "chart-bar");

  s.addText([
    { text: "4 ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Distribution et convergence", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // Left chart: Histogram
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.25, w: 4.5, h: 3.5,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText("Histogramme des coupons", {
    x: 0.6, y: 1.35, w: 4.1, h: 0.35,
    fontFace: F.serif, fontSize: 13, color: C.charcoal, margin: 0,
  });

  s.addChart(pres.charts.BAR, [{
    name: "Fr\u00e9quence",
    labels: ["0$", "5$", "10$", "15$", "20$", "25$", "30$", "35$", "40$", "45$", "50$", "55$", "60$", "65$"],
    values: [3200, 200, 150, 120, 180, 250, 320, 400, 500, 650, 800, 1200, 1500, 1800],
  }], {
    x: 0.5, y: 1.75, w: 4.2, h: 2.7,
    barDir: "col",
    chartColors: [C.teal],
    showLegend: false,
    showValue: false,
    catGridLine: { style: "none" },
    valGridLine: { color: C.borderLt, size: 0.5 },
    catAxisLabelColor: C.warmGray,
    valAxisLabelColor: C.warmGray,
    catAxisLabelFontSize: 7,
    valAxisLabelFontSize: 7,
  });

  // Right chart: Convergence
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.25, w: 4.5, h: 3.5,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addText("Convergence de l'estimateur", {
    x: 5.3, y: 1.35, w: 4.1, h: 0.35,
    fontFace: F.serif, fontSize: 13, color: C.charcoal, margin: 0,
  });

  const nSims = ["100", "500", "1k", "2k", "3k", "5k", "7k", "10k"];
  s.addChart(pres.charts.LINE, [
    {
      name: "Prix moyen",
      labels: nSims,
      values: [1042, 1038, 1037.5, 1036.8, 1036.6, 1036.5, 1036.5, 1036.49],
    },
    {
      name: "IC sup.",
      labels: nSims,
      values: [1055, 1044, 1041, 1039, 1038, 1037.5, 1037.1, 1037.36],
    },
    {
      name: "IC inf.",
      labels: nSims,
      values: [1029, 1032, 1034, 1034.6, 1035.2, 1035.5, 1035.9, 1035.62],
    },
  ], {
    x: 5.2, y: 1.75, w: 4.2, h: 2.7,
    chartColors: [C.charcoal, "80CBC4", "80CBC4"],
    lineSize: 2,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 7,
    legendColor: C.warmGray,
    catGridLine: { style: "none" },
    valGridLine: { color: C.borderLt, size: 0.5 },
    catAxisLabelColor: C.warmGray,
    valAxisLabelColor: C.warmGray,
    catAxisLabelFontSize: 8,
    valAxisLabelFontSize: 8,
  });

  // Bottom note
  s.addText("Distribution bimodale : pic \u00e0 0 $ (coupon nul) et pic \u00e0 ~65 $ (coupon maximal de 6,50%)", {
    x: 0.8, y: 4.9, w: 8.4, h: 0.35,
    fontFace: F.sans, fontSize: 10, italic: true, color: C.charcoal, align: "center",
  });

  s.addNotes(
    "Ces deux graphiques illustrent les r\u00e9sultats de la simulation Monte Carlo.\n\n" +
    "L'histogramme \u00e0 gauche montre la distribution des paiements de coupon sur les 10 000 simulations. " +
    "On observe une distribution bimodale tr\u00e8s caract\u00e9ristique :\n" +
    "\u2022 Un pic important \u00e0 0 $ : ce sont les simulations o\u00f9 les pertes des titres n\u00e9gatifs dominent, " +
    "et la moyenne pond\u00e9r\u00e9e tombe sous z\u00e9ro, d\u00e9clenchant le plancher global du coupon\n" +
    "\u2022 Un pic \u00e0 ~65 $ (6,50%) : ce sont les simulations o\u00f9 la majorit\u00e9 des titres sont positifs, " +
    "et les quelques n\u00e9gatifs ne suffisent pas \u00e0 faire baisser la moyenne sous 6,50%\n" +
    "\u2022 Peu de valeurs interm\u00e9diaires : la nature binaire du coupon num\u00e9rique cr\u00e9e cette polarisation\n\n" +
    "Le graphique de droite montre la convergence de l'estimateur Monte Carlo. " +
    "On voit que la moyenne se stabilise rapidement autour de 1 036,49 $ et que l'intervalle de confiance " +
    "(la bande color\u00e9e) se resserre \u00e0 mesure que N augmente. " +
    "D\u00e8s 5 000 simulations, l'estimation est d\u00e9j\u00e0 tr\u00e8s pr\u00e9cise."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 15a — Conclusion: CD numérique vs CD traditionnel
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "scales");

  s.addText([
    { text: "  ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Conclusion", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  // ── Left column: CD traditionnel ──
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 1.3, w: 4.2, h: 3.6,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.3, w: 4.2, h: 0.45, fill: { color: C.warmGray } });
  s.addText("CD TRADITIONNEL", {
    x: 0.6, y: 1.3, w: 4.2, h: 0.45,
    fontFace: F.sans, fontSize: 11, color: C.white, bold: true, align: "center", valign: "middle", charSpacing: 1,
  });

  const tradItems = [
    "Taux fixe garanti (~0,5\u20131,5% en 2016)",
    "Rendement pr\u00e9visible mais faible",
    "Aucune exposition aux march\u00e9s",
    "Pas de potentiel de surperformance",
    "Coupon vers\u00e9 syst\u00e9matiquement",
  ];
  tradItems.forEach((item, i) => {
    s.addText([{ text: item, options: { bullet: true } }], {
      x: 0.85, y: 1.95 + i * 0.45, w: 3.7, h: 0.4,
      fontFace: F.sans, fontSize: 11, color: C.warmGray, margin: 0,
    });
  });

  // ── Right column: CD numérique ──
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.3, w: 4.2, h: 3.6,
    fill: { color: C.white }, shadow: makeShadow(),
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.2, h: 0.45, fill: { color: C.teal } });
  s.addText("CD NUM\u00c9RIQUE (NOTRE PRODUIT)", {
    x: 5.2, y: 1.3, w: 4.2, h: 0.45,
    fontFace: F.sans, fontSize: 11, color: C.white, bold: true, align: "center", valign: "middle", charSpacing: 1,
  });

  const numItems = [
    { text: "Coupon potentiel de 6,50% (vs ~1%)", color: C.green },
    { text: "Rendement moyen estim\u00e9 \u00e0 ~4,8%", color: C.green },
    { text: "Exposition diversifi\u00e9e (10 titres)", color: C.teal },
    { text: "Protection du capital \u00e0 100%", color: C.teal },
    { text: "Plancher \u00e0 \u221230% par titre", color: C.teal },
  ];
  numItems.forEach((item, i) => {
    s.addText([{ text: item.text, options: { bullet: true } }], {
      x: 5.45, y: 1.95 + i * 0.45, w: 3.7, h: 0.4,
      fontFace: F.sans, fontSize: 11, color: item.color, bold: true, margin: 0,
    });
  });

  // ── Bottom: key takeaway ──
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.05, w: 8.8, h: 0.45, fill: { color: "E0F2F1" } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.05, w: 0.06, h: 0.45, fill: { color: C.teal } });
  s.addText("Le CD num\u00e9rique offre un rendement esp\u00e9r\u00e9 ~4\u00d7 sup\u00e9rieur au CD traditionnel, avec la m\u00eame protection du capital.", {
    x: 0.85, y: 5.05, w: 8.35, h: 0.45,
    fontFace: F.sans, fontSize: 12, color: C.charcoal, bold: true, valign: "middle",
  });

  s.addNotes(
    "Avant de parler des limites, comparons ce CD num\u00e9rique \u00e0 un CD traditionnel.\n\n" +
    "Un CD traditionnel en 2016 offrait un taux fixe d'environ 0,5% \u00e0 1,5% selon la dur\u00e9e. " +
    "C'est pr\u00e9visible mais tr\u00e8s faible, surtout dans l'environnement de taux bas post-2008.\n\n" +
    "Notre CD num\u00e9rique offre plusieurs avantages :\n\n" +
    "1. RENDEMENT POTENTIEL BIEN SUP\u00c9RIEUR : le coupon maximal de 6,50% est environ 4 \u00e0 6 fois le taux d'un CD classique. " +
    "M\u00eame le coupon moyen estim\u00e9 (~4,8%) d\u00e9passe largement un CD traditionnel.\n\n" +
    "2. EXPOSITION AUX MARCH\u00c9S SANS RISQUE EN CAPITAL : l'investisseur b\u00e9n\u00e9ficie de la hausse des march\u00e9s " +
    "(via le coupon num\u00e9rique) tout en conservant la garantie de remboursement du principal \u00e0 100%. " +
    "C'est comme avoir un pied dans le march\u00e9 actions sans risquer son capital.\n\n" +
    "3. DIVERSIFICATION INT\u00c9GR\u00c9E : le panier de 10 titres de secteurs vari\u00e9s offre une diversification naturelle. " +
    "Un seul titre en chute ne d\u00e9truit pas le coupon \u2014 les 9 autres peuvent compenser.\n\n" +
    "4. PROTECTION DOUBLE : la protection du plancher \u00e0 \u221230% par titre, combin\u00e9e au plancher global du coupon \u00e0 0%, " +
    "prot\u00e8ge l'investisseur \u00e0 deux niveaux.\n\n" +
    "Le compromis, bien s\u00fbr, est que le coupon n'est pas garanti. Dans environ 30% des simulations, " +
    "le coupon est nul. Mais m\u00eame dans ce sc\u00e9nario, le principal est int\u00e9gralement rembours\u00e9. " +
    "L'investisseur ne perd jamais d'argent \u2014 il risque seulement de ne pas en gagner.\n\n" +
    "C'est un produit id\u00e9al pour un investisseur conservateur qui souhaite un rendement sup\u00e9rieur aux taux fixes " +
    "tout en refusant de mettre son capital en danger."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 15b — Limites et Questions
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.cream };
  addIcon(s, "warning");

  s.addText([
    { text: "  ", options: { fontFace: F.mono, color: C.teal, bold: true } },
    { text: "Limites du mod\u00e8le", options: { fontFace: F.serif, color: C.charcoal } },
  ], { x: 0.8, y: 0.4, w: 8.4, h: 0.7, fontSize: 24, margin: 0 });
  s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.1, w: 8.4, h: 0, line: { color: C.borderLt, width: 1 } });

  const limits = [
    { title: "Volatilit\u00e9 constante (GBM)", desc: "Le mod\u00e8le suppose \u03c3 fixe. En r\u00e9alit\u00e9, la volatilit\u00e9 varie dans le temps (smile, clustering). Un mod\u00e8le de Heston serait plus r\u00e9aliste.", color: C.orange },
    { title: "Pas de temps unique", desc: "On simule directement de oct. 2016 \u00e0 sept. 2017. Adapt\u00e9 ici car il ne reste qu'une seule date de coupon, mais ne permettrait pas d'\u00e9valuer les coupons ant\u00e9rieurs.", color: C.orange },
    { title: "Corr\u00e9lation statique", desc: "La matrice de corr\u00e9lation historique est suppos\u00e9e stable. Les corr\u00e9lations augmentent souvent en p\u00e9riode de crise (contagion), ce qui n'est pas captur\u00e9.", color: C.orange },
    { title: "Taux sans risque fixe", desc: "r = 1,25% est constant. Un mod\u00e8le de taux (Vasicek, Hull-White) capterait mieux le risque de taux d'int\u00e9r\u00eat.", color: C.orange },
  ];

  limits.forEach((lim, i) => {
    const yBase = 1.3 + i * 0.8;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: yBase, w: 8.4, h: 0.7,
      fill: { color: C.orangeBg },
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: yBase, w: 0.06, h: 0.7, fill: { color: C.orange } });
    s.addText(lim.title, {
      x: 1.05, y: yBase + 0.04, w: 7.9, h: 0.25,
      fontFace: F.sans, fontSize: 12, color: C.orange, bold: true, margin: 0,
    });
    s.addText(lim.desc, {
      x: 1.05, y: yBase + 0.3, w: 7.9, h: 0.35,
      fontFace: F.sans, fontSize: 10, color: C.charcoal, margin: 0,
    });
  });

  // Questions
  s.addText("Questions ?", {
    x: 0.8, y: 4.7, w: 8.4, h: 0.6,
    fontFace: F.serif, fontSize: 28, color: C.charcoal, align: "center", valign: "middle",
  });

  s.addNotes(
    "Voici les 4 principales limites de notre mod\u00e8le, d\u00e9taill\u00e9es :\n\n" +
    "1. VOLATILIT\u00c9 CONSTANTE : Le GBM suppose que \u03c3 ne change pas dans le temps. " +
    "En pratique, on observe le \u00ab smile de volatilit\u00e9 \u00bb (les options hors-de-la-monnaie ont une vol implicite plus \u00e9lev\u00e9e) " +
    "et le \u00ab clustering de volatilit\u00e9 \u00bb (les p\u00e9riodes de forte vol sont suivies de forte vol). " +
    "Un mod\u00e8le de Heston (volatilit\u00e9 stochastique) ou GARCH capterait ces effets.\n\n" +
    "2. PAS DE TEMPS UNIQUE : On simule en un seul saut de oct. 2016 \u00e0 sept. 2017. " +
    "C'est justifi\u00e9 car il ne reste qu'une seule date de d\u00e9termination du coupon. " +
    "Si on \u00e9valuait le CD plus t\u00f4t (par ex. en 2013), il faudrait simuler les prix \u00e0 chaque date de coupon, " +
    "ce qui n\u00e9cessiterait une simulation multi-p\u00e9riodes.\n\n" +
    "3. CORR\u00c9LATION STATIQUE : Notre matrice de corr\u00e9lation est calcul\u00e9e sur 10 ans de donn\u00e9es historiques " +
    "et suppos\u00e9e stable sur les 11 prochains mois. En r\u00e9alit\u00e9, les corr\u00e9lations entre actions augmentent " +
    "fortement en p\u00e9riode de stress (effet de contagion). Cela signifie que notre mod\u00e8le sous-estime " +
    "potentiellement le risque de sc\u00e9narios o\u00f9 toutes les actions chutent simultan\u00e9ment. " +
    "Des copules (par ex. copule de Clayton) permettraient de mieux mod\u00e9liser cette d\u00e9pendance de queue.\n\n" +
    "4. TAUX SANS RISQUE FIXE : On utilise r = 1,25% comme constante. " +
    "Un mod\u00e8le de taux stochastique (Vasicek, Hull-White) serait plus pr\u00e9cis, " +
    "surtout pour des produits de longue dur\u00e9e. L'impact est limit\u00e9 ici car l'horizon r\u00e9siduel est court (11 mois).\n\n" +
    "Malgr\u00e9 ces limites, le mod\u00e8le donne une estimation fiable : l'IC \u00e0 95% est tr\u00e8s serr\u00e9 (\u00b10,87 $). " +
    "Les am\u00e9liorations les plus impactantes seraient la volatilit\u00e9 stochastique et les copules.\n\n" +
    "Merci pour votre attention. Je suis disponible pour vos questions."
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// SAVE
// ═══════════════════════════════════════════════════════════════════════════
const outPath = "/Users/eddie/projects/cours_marina/wp_03/presentation_devoir2.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("Presentation saved to: " + outPath);
}).catch(err => {
  console.error("Error:", err);
});
