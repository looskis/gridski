const ALIGN = { left: "horizontal align left", center: "horizontal align center", right: "horizontal align right", general: "horizontal align general" };
const EDGES = { top: "edge top", bottom: "edge bottom", left: "edge left", right: "edge right", inside_horizontal: "inside horizontal", inside_vertical: "inside vertical" };
const LINES = {
  thin: ["continuous", "border weight thin"],
  medium: ["continuous", "border weight medium"],
  thick: ["continuous", "border weight thick"],
  double: ["double", "border weight thick"],
  dashed: ["dash", "border weight thin"],
  none: ["line style none", null],
};

respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const rg = resolveRange(sh, ARGS.range);
  const applied = [];
  const has = (k) => ARGS[k] !== null && ARGS[k] !== undefined;

  if (has("font_color")) (rg.fontObject.color = rgbFromHex(ARGS.font_color)), applied.push("font_color");
  if (has("bold")) (rg.fontObject.bold = ARGS.bold), applied.push("bold");
  if (has("italic")) (rg.fontObject.italic = ARGS.italic), applied.push("italic");
  if (has("fill")) {
    if (ARGS.fill === "none") rg.interiorObject.colorIndex = "color index none";
    else rg.interiorObject.color = rgbFromHex(ARGS.fill);
    applied.push("fill");
  }
  if (has("number_format")) (rg.numberFormat = ARGS.number_format), applied.push("number_format");
  if (has("horizontal_alignment")) {
    if (!ALIGN[ARGS.horizontal_alignment]) throw new GridskiError("INVALID", "horizontal_alignment is left, center, right, or general.");
    rg.horizontalAlignment = ALIGN[ARGS.horizontal_alignment];
    applied.push("horizontal_alignment");
  }
  if (has("indent")) (rg.indentLevel = ARGS.indent), applied.push("indent");
  if (has("column_width")) (rg.columnWidth = ARGS.column_width), applied.push("column_width");
  if (has("borders")) {
    for (const [edge, style] of Object.entries(ARGS.borders).filter(([, v]) => v !== null)) {
      if (!EDGES[edge]) throw new GridskiError("INVALID", `Unknown border edge "${edge}".`);
      if (!LINES[style]) throw new GridskiError("INVALID", `Unknown border style "${style}".`);
      const border = excel.getBorder(rg, { whichBorder: EDGES[edge] });
      border.lineStyle = LINES[style][0];
      if (LINES[style][1]) border.weight = LINES[style][1];
    }
    applied.push("borders");
  }
  return { workbook: wb.name(), sheet: sh.name(), address: rg.getAddress().replace(/\$/g, ""), applied };
});
