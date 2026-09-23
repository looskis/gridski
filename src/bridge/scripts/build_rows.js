// Write a block of time-series rows in one round trip: the label…first-period grid,
// then every later period as a copy of each row's first cell (R1C1), then formats.
// ARGS: workbook, sheet, overwrite, and a plan from src/rows.rs.
const LINE = ["continuous", "border weight thin"];

respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const target = sh.ranges[ARGS.target];

  if (!ARGS.overwrite) {
    const [t, l] = bounds(ARGS.target);
    const filled = [];
    grid(target.formula()).forEach((row, i) =>
      row.forEach((f, j) => {
        if (f !== "" && f !== null && filled.length < 5) filled.push(colLetters(l + j) + (t + i));
      })
    );
    if (filled.length) {
      throw new GridskiError("INVALID", `${ARGS.target} is not empty (e.g. ${filled.join(", ")}); pass overwrite: true to replace it.`);
    }
  }

  const head = ARGS.head.map((row) => row.map((c) => (c === null ? "" : c)));
  const [ht, hl] = bounds(ARGS.head_start);
  const headRange = sh.ranges[a1(ht, hl, ht + head.length - 1, hl + head[0].length - 1)];
  headRange.formula = head.length === 1 && head[0].length === 1 ? head[0][0] : head;

  if (ARGS.fill) {
    const [ft, fl, fb, fr] = bounds(ARGS.fill);
    const first = grid(sh.ranges[a1(ft, ARGS.first_col, fb, ARGS.first_col)].formulaR1c1());
    const cols = fr - fl + 1;
    const data = first.map((row) => new Array(cols).fill(row[0] === null ? "" : row[0]));
    sh.ranges[ARGS.fill].formulaR1c1 = data.length === 1 && cols === 1 ? data[0][0] : data;
  }

  for (const op of ARGS.formats) {
    const rg = sh.ranges[op.range];
    if (op.number_format) rg.numberFormat = op.number_format;
    if (op.bold) rg.fontObject.bold = true;
    if (op.font_color) rg.fontObject.color = rgbFromHex(op.font_color);
    if (op.top_border) {
      const border = excel.getBorder(rg, { whichBorder: "edge top" });
      border.lineStyle = LINE[0];
      border.weight = LINE[1];
    }
  }

  return { workbook: wb.name(), sheet: sh.name(), address: ARGS.target, rows: ARGS.keys, formats: ARGS.formats.length };
});
