respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const name = ARGS.name;
  const existing = (wb.namedItems.name() || []).includes(name) ? wb.namedItems[name] : null;
  const previous = existing ? existing.referenceLocal() : null;
  const itemRef = `named item ${asq(name)} of workbook ${asq(wb.name())}`;

  if (ARGS.refers_to === null || ARGS.refers_to === undefined) {
    if (!existing) throw new GridskiError("NOT_FOUND", `No name "${name}" in "${wb.name()}".`);
    excelScript(`delete ${itemRef}`);
    return { workbook: wb.name(), name, refers_to: null, previous };
  }
  const refersTo = ARGS.refers_to.startsWith("=") ? ARGS.refers_to : "=" + ARGS.refers_to;
  if (existing) {
    existing.referenceLocal = refersTo;
  } else {
    excelScript(
      `make new named item at workbook ${asq(wb.name())} with properties {name:${asq(name)}, references:${asq(refersTo)}}`
    );
  }
  return { workbook: wb.name(), name, refers_to: wb.namedItems[name].referenceLocal(), previous };
});
