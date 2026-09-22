const FORMATS = { xlsx: "Excel XML file format", xlsm: "macro enabled XML file format", xlsb: "Excel binary file format" };

respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const current = wb.fullName();
  const path = ARGS.path || current;
  if (!path.startsWith("/")) {
    throw new GridskiError("INVALID", `"${wb.name()}" has never been saved; pass a path ending in .xlsx.`);
  }
  // Saving in place is only safe inside the workspace too; elsewhere Excel may prompt.
  if (!path.startsWith(ARGS.root + "/")) {
    throw new GridskiError("INVALID", `${path} is outside the workspace folder ${ARGS.root}; pass a path inside it to save a copy there.`);
  }
  const ext = path.split(".").pop().toLowerCase();
  if (!FORMATS[ext]) throw new GridskiError("INVALID", "Save as .xlsx, .xlsm, or .xlsb.");

  const before = fileMtime(path);
  if (path === current) {
    excelScript(`save workbook ${asq(wb.name())}`);
  } else {
    if (before !== null && !ARGS.overwrite) {
      throw new GridskiError("INVALID", `${path} already exists; pass overwrite: true to replace it.`);
    }
    excelScript(`save workbook as workbook ${asq(wb.name())} filename ${asq(path)} file format ${FORMATS[ext]}`);
  }
  // Folders Excel hasn't been granted access to fail the same silent way; verify the write.
  const after = fileMtime(path);
  if (after === null || (before !== null && after === before)) {
    throw new GridskiError(
      "UNSUPPORTED",
      `Excel reported saving but ${path} was not written; Excel may lack access to that folder. ` +
        "Grant Excel access to the workspace folder when it asks, then save again."
    );
  }
  const saved = excel.workbooks[path.split("/").pop()];
  return { workbook: saved.name(), path: saved.fullName(), saved: saved.saved() };
});
