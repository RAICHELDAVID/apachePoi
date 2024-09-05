<!---<cfscript>
    // Import necessary Apache POI classes
    workbook = createObject("java", "org.apache.poi.xssf.usermodel.XSSFWorkbook").init();
    theFile="Total sections.xlsx";
    sheet = workbook.createSheet("Total sections sheet");

    // Create cell styles
    mainHeadingStyle = workbook.createCellStyle();
    contentStyle = workbook.createCellStyle();

    // Set font styles
    font = workbook.createFont();
    font.setFontName("Arial Narrow");
    font.setFontHeightInPoints(12);
    font.setBold(true);
    font.setUnderline(1); // Font.UNDERLINE_SINGLE

    mainHeadingStyle.setFont(font);
    mainHeadingStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").RIGHT);
    mainHeadingStyle.setFillForegroundColor(createObject("java", "org.apache.poi.ss.usermodel.IndexedColors").YELLOW.getIndex());
    mainHeadingStyle.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").SOLID_FOREGROUND);

    // Configure content style
    fontContent = workbook.createFont();
    fontContent.setFontName("Arial Narrow");
    fontContent.setFontHeightInPoints(12);

    contentStyle.setFont(fontContent);
    contentStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").RIGHT);

    // Row and column indices
    rowIdx = 0;
    colIdx = 0;

    // Create rows and cells
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("TOTAL SECTIONS 1-2:");
    cell.setCellStyle(mainHeadingStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Sub Total:");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total DFI % (No Spoils):");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total Cost:");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    // Write the workbook to a file
    fileOutputStream = createObject("java", "java.io.FileOutputStream").init("Total_sections.xlsx");
    workbook.write(fileOutputStream);
    fileOutputStream.close();
    workbook.close();

    // Notify that the file was created successfully
    writeOutput("Spreadsheet created successfully with custom formatting!");
</cfscript>

<!--- Set headers to prompt file download --->
<cfheader name="Content-Disposition" value="inline; filename=#theFile#">
<cfcontent type="application/vnd.ms-excel" variable="#SpreadSheetReadBinary(sheet)#">
--->

<cfscript>
    // Import necessary Apache POI classes
    workbook = createObject("java", "org.apache.poi.xssf.usermodel.XSSFWorkbook").init();
    sheet = workbook.createSheet("Total sections sheet");

    // Create cell styles
    mainHeadingStyle = workbook.createCellStyle();
    contentStyle = workbook.createCellStyle();

    // Set font styles
    font = workbook.createFont();
    font.setFontName("Arial Narrow");
    font.setFontHeightInPoints(12);
    font.setBold(true);
    font.setUnderline(1); // Font.UNDERLINE_SINGLE

    mainHeadingStyle.setFont(font);
    //mainHeadingStyle.setFillForegroundColor(createObject("java", "org.apache.poi.ss.usermodel.IndexedColors").rgba(219,219,219,255).getIndex());
    greyColor = createObject("java", "org.apache.poi.xssf.usermodel.XSSFColor").init(createObject("java", "java.awt.Color").init(219, 219, 219));
    mainHeadingStyle.setFillForegroundColor(greyColor);

    mainHeadingStyle.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").SOLID_FOREGROUND);

    // Configure content style
    fontContent = workbook.createFont();
    fontContent.setFontName("Arial Narrow");
    fontContent.setFontHeightInPoints(12);

    contentStyle.setFont(fontContent);
    contentStyle.setAlignment(createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment").RIGHT);

    // Row and column indices
    rowIdx = 0;
    colIdx = 0;

    // Create rows and cells
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("TOTAL SECTIONS 1-2:");
    cell.setCellStyle(mainHeadingStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Sub Total:");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total DFI % (No Spoils):");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    rowIdx++;
    row = sheet.createRow(rowIdx);
    cell = row.createCell(colIdx);
    cell.setCellValue("Total Cost:");
    cell.setCellStyle(contentStyle);

    sheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(rowIdx, rowIdx, colIdx, colIdx+1));

    // Write the workbook to a ByteArrayOutputStream
    baos = createObject("java", "java.io.ByteArrayOutputStream").init();
    workbook.write(baos);
    workbook.close();

    // Set headers and content for download
    theFile = "Total_sections.xlsx";
    cfheader(name="Content-Disposition", value="attachment; filename=#theFile#");
    cfcontent(type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", variable="#baos.toByteArray()#");
</cfscript>
