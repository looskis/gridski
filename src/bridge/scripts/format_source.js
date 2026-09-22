respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, ARGS);
  const range = ARGS.range ? resolveRange(sh, ARGS.range) : sh.usedRange;
  return { workbook: wb.name(), sheet: sh.name(), address: range.getAddress(), path: wb.fullName(), saved: wb.saved() };
});
