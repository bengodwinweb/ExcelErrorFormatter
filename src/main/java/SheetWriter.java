import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetWriter {
    public static final int NUM_HEADER_ROWS = 3;
    public static final int DATE_COL = 0;
    public static final int FILE_COL = Main.FILE_COL;
    public static final int NO_06A_COL = 2;
    public static final int _06A_COL = 3;
    public static final int _005_COL = 6;
    public static final int _031_COL = 9;
    public static final int _065_COL = 12;
    public static final int _066_COL = 15;
    public static final int _067_COL = 18;
    public static final int NUM_COLS = Main.BunoErrors.get(Main.BunoErrors.size() - 2).getEndCol() + 1;
    public static IndexedColors headerColor = IndexedColors.PALE_BLUE;

    private Workbook wb;
    private Sheet sheet;
    private List<CustomRowData> rowData;

    private Row firstRow;
    private Row secondRow;
    private Row thirdRow;
    private Row footerRow;

    public SheetWriter(Workbook wb, String sheetName, List<CustomRowData> rowData) {
        this.wb = wb;
        this.sheet = wb.createSheet("BUNO - " + sheetName);
        this.rowData = rowData;

        firstRow = sheet.createRow(0);
        secondRow = sheet.createRow(1);
        thirdRow = sheet.createRow(2);
        footerRow = sheet.createRow(NUM_HEADER_ROWS + rowData.size());
    }

    public void makeSheet() {
        makeHeaderAndFooter();
        makeDataRows();
        makeBorders();
    }

    private void makeHeaderAndFooter() {
        // Font for the header and footer cells - set bold to true
        Font headerFont = wb.createFont();
        headerFont.setBold(true);

        // style for all header and footer cells that are not for a specific error - background is PALE_BLUE
        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(headerColor.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // set the style for header and footer cells related to each error - get background color from the error
        for (BunoError e : Main.BunoErrors) {
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setFont(headerFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(e.getColor().getIndex());
            e.setCellStyle(cellStyle);
        }

        // create all cells in the first row and set the style to header style (there are no error specific cells in the first row)
        for (int i = 0; i < NUM_COLS; i++) firstRow.createCell(i).setCellStyle(headerStyle);

        // styling for second, third, and footer rows
        List<Row> styledRows = Arrays.asList(secondRow, thirdRow, footerRow);
        // set the first two cells to header style
        for (int i = 0; i <= FILE_COL; i++) {
            for (Row row : styledRows) row.createCell(i).setCellStyle(headerStyle);
        }
        // go through the errors and set the style for each col within the error's range to that error's style
        for (BunoError e : Main.BunoErrors) {
            for (int i = e.getStartCol(); i <= e.getEndCol(); i++) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(e.getCellStyle());
            }
        }

        // set the values for cells that are not variable
        firstRow.getCell(DATE_COL).setCellValue("Date");
        firstRow.getCell(FILE_COL).setCellValue("File");
        firstRow.getCell(FILE_COL + 1).setCellValue("MSP Codes");
        footerRow.getCell(DATE_COL).setCellValue("TOTAL:");

        // set the values for the Error Code headers in the second row
        for (BunoError e : Main.BunoErrors) secondRow.getCell(e.getStartCol()).setCellValue(e.getCode());

        // set the values for the third row under the error code header - Pre-Flight then In-Flight then Post-Flight for each (except No_06A)
        for (int i = Main.BunoErrors.get(0).getStartCol(); i < NUM_COLS; i++) {
            if (i % 3 == (Main.BunoErrors.get(0).getStartCol() % 3)) thirdRow.getCell(i).setCellValue("Pre-Flight");
            if (i % 3 == ((Main.BunoErrors.get(0).getStartCol() + 1) % 3))
                thirdRow.getCell(i).setCellValue("In-Flight");
            if (i % 3 == ((Main.BunoErrors.get(0).getStartCol() + 2) % 3))
                thirdRow.getCell(i).setCellValue("Post-Flight");
        }

        // create merged regions in first and footer rows
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, DATE_COL, DATE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, FILE_COL, FILE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), FILE_COL + 1, NUM_COLS - 1));
        sheet.addMergedRegion(new CellRangeAddress(footerRow.getRowNum(), footerRow.getRowNum(), DATE_COL, FILE_COL));

        // create merged regions in second row for each error over the span of its cols
        for (int i = 0; i < Main.BunoErrors.size() - 1; i++) {
            BunoError e = Main.BunoErrors.get(i);
            sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), e.getStartCol(), e.getEndCol()));
        }
        // merged region for No_06A error in its col from second to third row
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), Main.BunoErrors.get(Main.BunoErrors.size() - 1).getStartCol(), Main.BunoErrors.get(Main.BunoErrors.size() - 1).getEndCol()));
    }

    private void makeBorders() {
        PropertyTemplate pt = new PropertyTemplate();

        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();

        // Draw thick border around header and fill in light borders
        CellRangeAddress header = new CellRangeAddress(firstRow.getRowNum(), thirdRow.getRowNum(), DATE_COL, NUM_COLS - 1);
        cellRangeAddresses.add(header);
        cellRangeAddresses.addAll(createRange(DATE_COL, DATE_COL));
        cellRangeAddresses.addAll(createRange(FILE_COL, FILE_COL));
        for (BunoError e : Main.BunoErrors) cellRangeAddresses.addAll(createRange(e));

        for (CellRangeAddress range : cellRangeAddresses) {
            pt.drawBorders(range, BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
            pt.drawBorders(range, BorderStyle.THIN, BorderExtent.INSIDE);
        }

        pt.applyBorders(sheet);
    }

    private void makeDataRows() {
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("m/d/yy"));

        CellStyle lightCountCellStyle = wb.createCellStyle();
        lightCountCellStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle darkCountCellStyle = wb.createCellStyle();
        darkCountCellStyle.setAlignment(HorizontalAlignment.CENTER);
        darkCountCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        darkCountCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < rowData.size(); i++) {
            Row row = sheet.createRow(i + NUM_HEADER_ROWS);
            addRowContents(row, rowData.get(i));
            row.getCell(DATE_COL).setCellStyle(dateStyle);
            for (int j = NO_06A_COL; j < NUM_COLS; j++) {
                if (j < _005_COL) row.getCell(j).setCellStyle(darkCountCellStyle);
                else if (j < _031_COL) row.getCell(j).setCellStyle(lightCountCellStyle);
                else if (j < _065_COL) row.getCell(j).setCellStyle(darkCountCellStyle);
                else if (j < _066_COL) row.getCell(j).setCellStyle(lightCountCellStyle);
                else if (j < _067_COL) row.getCell(j).setCellStyle(darkCountCellStyle);
                else row.getCell(j).setCellStyle(lightCountCellStyle);
            }
        }

        sheet.autoSizeColumn(DATE_COL);
        sheet.autoSizeColumn(FILE_COL);

        for (int i = NO_06A_COL; i < NUM_COLS; i++) {
            char colChar = (char) ('A' + i);
            String sumColFormula = "SUM(" + colChar + (NUM_HEADER_ROWS + 1) + ":" + colChar + (footerRow.getRowNum()) + ")";
            footerRow.getCell(i).setCellFormula(sumColFormula);
        }
    }

    private void addRowContents(Row row, CustomRowData rowData) {
        row.createCell(DATE_COL).setCellValue(rowData.getDate());
        row.createCell(FILE_COL).setCellValue(rowData.getFileName());
        boolean[] events = rowData.getEventArray();
        for (int i = 0; i < events.length; i++) {
            if (events[i]) row.createCell(i + NO_06A_COL).setCellValue(1);
            else row.createCell(i + NO_06A_COL);
        }
    }

    private List<CellRangeAddress> createRange(int startRow, int endRow, int startCol, int endCol) {
        CellRangeAddress header = new CellRangeAddress(startRow, endRow, startCol, endCol);
        CellRangeAddress footer = new CellRangeAddress(footerRow.getRowNum(), footerRow.getRowNum(), startCol, endCol);
        CellRangeAddress body = new CellRangeAddress(NUM_HEADER_ROWS, footerRow.getRowNum() - 1, startCol, endCol);
        return Arrays.asList(header, footer, body);
    }

    private List<CellRangeAddress> createRange(int startCol, int endCol) {
        return createRange(firstRow.getRowNum(), thirdRow.getRowNum(), startCol, endCol);
    }

    private List<CellRangeAddress> createRange(BunoError e) {
        return createRange(secondRow.getRowNum(), thirdRow.getRowNum(), e.getStartCol(), e.getEndCol());
    }
}
