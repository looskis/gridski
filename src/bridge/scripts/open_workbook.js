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
  excel.openWorkbook({ workbookFileName: path, updateLinks: "do not update links", readOnly: !!ARGS.read_only });
  return workbookInfo(excel, excel.workbooks[name]);
});
