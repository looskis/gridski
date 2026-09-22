respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const wbRef = `workbook ${asq(wb.name())}`;
  const names = () => wb.worksheets.name();
  const need = (sheet) => {
    if (!sheet) throw new GridskiError("INVALID", `"${ARGS.action}" needs a sheet.`);
    if (!names().includes(sheet)) throw new GridskiError("NOT_FOUND", `No sheet named "${sheet}" in "${wb.name()}".`);
    return `worksheet ${asq(sheet)} of ${wbRef}`;
  };
  const position = () => {
    if (ARGS.before) return `before ${need(ARGS.before)}`;
    if (ARGS.after) return `after ${need(ARGS.after)}`;
    return `after worksheet ${asq(names().slice(-1)[0])} of ${wbRef}`;
  };
  const freeName = (name) => {
    if (names().some((n) => n.toLowerCase() === name.toLowerCase())) {
      throw new GridskiError("INVALID", `A sheet named "${name}" already exists.`);
    }
  };

  let sheet = ARGS.sheet;
  switch (ARGS.action) {
    case "add": {
      if (ARGS.name) freeName(ARGS.name);
      // Name it through its new name, not `s`: the reference goes stale once renamed.
      sheet = excelScript(`set s to make new worksheet at ${position()}\nreturn name of s`);
      if (ARGS.name) {
        wb.worksheets[sheet].name = ARGS.name;
        sheet = ARGS.name;
      }
      break;
    }
    case "rename": {
      need(sheet);
      if (!ARGS.name) throw new GridskiError("INVALID", "rename needs name.");
      if (ARGS.name.toLowerCase() !== sheet.toLowerCase()) freeName(ARGS.name);
      wb.worksheets[sheet].name = ARGS.name;
      sheet = ARGS.name;
      break;
    }
    case "move": {
      const target = need(sheet);
      if (!ARGS.before && !ARGS.after) throw new GridskiError("INVALID", "move needs before or after.");
      excelScript(`move ${target} to ${position()}`);
      break;
    }
    case "delete": {
      const target = need(sheet);
      if (names().length === 1) throw new GridskiError("UNSUPPORTED", "A workbook must keep at least one sheet.");
      excelScript(`delete ${target}`);
      break;
    }
    default:
      throw new GridskiError("INVALID", `Unknown action "${ARGS.action}".`);
  }
  return { workbook: wb.name(), sheet, sheets: names() };
});
