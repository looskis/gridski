respond(() => {
  const excel = excelApp();
  if (!ARGS.path) return workbookInfo(excel, excel.make({ new: "workbook" }));

  const path = ARGS.path;
  const name = path.split("/").pop();
  const open = excel.workbooks().find((wb) => wb.name() === name);
  if (open) {
    // Excel can't hold two workbooks with the same file name.
    if (open.fullName() === path) return workbookInfo(excel, open);
    throw new GridskiError("UNSUPPORTED", `A different "${name}" is already open (${open.fullName()}); close it first.`);
  }
  if (fileMtime(path) === null) throw new GridskiError("NOT_FOUND", `No file at ${path}.`);
  // Open with macros force-disabled: Excel's "enable macros?" prompt is modal and would
  // block every later call. Links are not updated, which would prompt too.
  applescript(
    `tell application "Microsoft Excel"\nset prev to automation security\n` +
      `set automation security to msoAutomationSecurityForceDisable\ntry\n` +
      `open workbook workbook file name ${asq(path)} update links do not update links read only ${!!ARGS.read_only}\n` +
      `on error m number n\nset automation security to prev\nerror m number n\nend try\n` +
      `set automation security to prev\nend tell`
  );
  return workbookInfo(excel, excel.workbooks[name]);
});
