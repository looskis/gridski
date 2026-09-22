respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const range = ARGS.range ? resolveRange(sh, ARGS.range) : sh.usedRange;
  return readRange(wb, sh, range, ARGS.max_cells, ARGS.include_formulas);
});
