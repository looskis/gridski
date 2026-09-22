respond(() => {
  const excel = excelApp();
  const wb = resolveWorkbook(excel, ARGS.workbook);
  const items = wb.namedItems;
  const names = items.name() || [];
  if (names.length === 0) return { workbook: wb.name(), names: [] };
  const refs = items.referenceLocal();
  const visible = items.visible();
  return {
    workbook: wb.name(),
    names: names
      .map((name, i) => ({ name, refers_to: refs[i], hidden: !visible[i] }))
      // Excel's own bookkeeping names (_xlfn., _xleta., _xlnm.) aren't the model's.
      .filter((n) => !n.name.startsWith("_xl")),
  };
});
