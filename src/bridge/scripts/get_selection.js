respond(() => {
  const excel = excelApp();
  const { wb, sh } = resolveSheet(excel, {});
  let address;
  try {
    address = excel.selection.getAddress();
  } catch (e) {
    throw new GridskiError("UNSUPPORTED", "The current selection is not a cell range (a chart or shape may be selected).");
  }
  // For multi-area selections, return the first area's contents and list every area.
  const areas = address.split(",");
  const data = readRange(wb, sh, sh.ranges[areas[0]], ARGS.max_cells, ARGS.include_formulas);
  if (areas.length > 1) data.areas = areas;
  return data;
});
