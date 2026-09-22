respond(() => {
  const excel = excelApp();
  return excel
    .workbooks()
    // Skip hidden workbooks such as PERSONAL.XLSB.
    .filter((wb) => wb.windows().some((w) => w.visible()))
    .map((wb) => workbookInfo(excel, wb));
});
