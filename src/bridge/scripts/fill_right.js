// Like Excel's Fill Right (Ctrl+R): copy each row's first cell across the rest of the row,
// adjusting relative references. Done in R1C1 so one assignment fills the whole block.
// Formulas and values only; formats are left as they are.
const MAX_PREVIOUS = 500;

respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const rg = resolveRange(sh, ARGS.range);
  const [t, l, b, r] = bounds(rg.getAddress());
  if (r === l) throw new GridskiError("INVALID", "range must be at least two columns wide; its first column is copied across the rest.");
  const source = grid(sh.ranges[a1(t, l, b, l)].formulaR1c1());
  const target = sh.ranges[a1(t, l + 1, b, r)];
  const before = grid(target.formula());
  const overwritten = {};
  let count = 0;
  before.forEach((row, i) =>
    row.forEach((f, j) => {
      if (f === "" || f === null) return;
      count += 1;
      if (count <= MAX_PREVIOUS) overwritten[colLetters(l + 1 + j) + (t + i)] = f;
    })
  );
  const cols = r - l;
  const data = source.map((row) => new Array(cols).fill(row[0] === null ? "" : row[0]));
  target.formulaR1c1 = data.length === 1 && cols === 1 ? data[0][0] : data;
  return {
    workbook: wb.name(),
    sheet: sh.name(),
    address: target.getAddress(),
    overwritten_count: count,
    previous_formulas: count <= MAX_PREVIOUS ? overwritten : null,
  };
});
