respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const name = wb.name();
  if (ARGS.save && !wb.fullName().startsWith("/")) {
    throw new GridskiError("INVALID", `"${name}" has never been saved; save_workbook it with a path first.`);
  }
  const path = wb.fullName();
  excel.close(wb, { saving: ARGS.save ? "yes" : "no" });
  // Excel can accept the close and still keep the workbook (busy finishing a save, or a
  // dialog up), so confirm it is gone before reporting success.
  for (let i = 0; i < 20; i++) {
    if (!excel.workbooks().some((w) => w.fullName() === path)) {
      return { closed: name, saved: ARGS.save };
    }
    delay(0.25);
  }
  throw new GridskiError(
    "UNSUPPORTED",
    `Excel did not close "${name}"; it is still open, probably because Excel is busy or showing a dialog. ` +
      "Ask the user to check Excel, then try again."
  );
});
