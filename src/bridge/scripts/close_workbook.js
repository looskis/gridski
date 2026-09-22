respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const name = wb.name();
  if (ARGS.save && !wb.fullName().startsWith("/")) {
    throw new GridskiError("INVALID", `"${name}" has never been saved; save_workbook it with a path first.`);
  }
  excel.close(wb, { saving: ARGS.save ? "yes" : "no" });
  return { closed: name, saved: ARGS.save };
});
