respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const range = ARGS.range ? resolveRange(sh, ARGS.range) : sh.usedRange;
  const totalRows = range.rows.length;
  const totalCols = range.columns.length;
  const cols = Math.min(totalCols, ARGS.max_cells);
  const rows = Math.min(totalRows, Math.max(1, Math.floor(ARGS.max_cells / cols)));
  const r = rows === totalRows && cols === totalCols ? range : range.getResize({ rowSize: rows, columnSize: cols });
  return {
    workbook: wb.name(),
    sheet: sh.name(),
    address: r.getAddress(),
    total_rows: totalRows,
    total_cols: totalCols,
    formulas: grid(r.formula()),
    r1c1: grid(r.formulaR1c1()),
  };
});
