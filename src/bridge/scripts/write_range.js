respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const values = ARGS.values;
  const rows = values.length;
  const cols = values[0].length;
  const target = resolveRange(sh, ARGS.start).getResize({ rowSize: rows, columnSize: cols });
  // Snapshot the prior contents so the caller can restore them: writing formulas back
  // reproduces constants too, since formula() returns "3" for a constant 3.
  const previous = grid(target.formula());
  const data = values.map((row) => row.map((c) => (c === null ? "" : c)));
  target.formula = rows === 1 && cols === 1 ? data[0][0] : data;
  return {
    workbook: wb.name(),
    sheet: sh.name(),
    address: target.getAddress(),
    previous_formulas: previous,
  };
});
