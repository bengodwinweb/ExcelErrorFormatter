import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

import java.util.Arrays;
import java.util.List;

public class SheetWriter {
    public static final int FIRST_ROW = 9;
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
    public static final int NUM_COLS = 21;
    public static IndexedColors headerColor = IndexedColors.PALE_BLUE;
    public static IndexedColors No_06A_Color = IndexedColors.PLUM;
    public static IndexedColors _06A_Color = IndexedColors.CORAL;
    public static IndexedColors _005_Color = IndexedColors.LIGHT_GREEN;
    public static IndexedColors _031_Color = IndexedColors.LIGHT_CORNFLOWER_BLUE;
    public static IndexedColors _065_Color = IndexedColors.LIGHT_YELLOW;
    public static IndexedColors _066_Color = IndexedColors.LIGHT_ORANGE;
    public static IndexedColors _067_Color = IndexedColors.LIGHT_TURQUOISE;

    private Workbook wb;
    private Sheet sheet;
    private List<CustomRowData> rowData;

    private Row firstRow;
    private Row secondRow;
    private Row thirdRow;
    private Row totalsRow;

    public SheetWriter(Workbook wb, String sheetName, List<CustomRowData> rowData) {
        this.wb = wb;
        this.sheet = wb.createSheet("BUNO - " + sheetName);
        this.rowData = rowData;

        firstRow = sheet.createRow(0);
        secondRow = sheet.createRow(1);
        thirdRow = sheet.createRow(2);
        totalsRow = sheet.createRow(NUM_HEADER_ROWS + rowData.size());
    }

    public void makeSheet() {
        makeHeaderAndFooter();
        makeDataRows();
        makeBorders();
    }

    private void makeHeaderAndFooter() {
        Font headerFont = wb.createFont();
        headerFont.setBold(true);

        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(headerColor.getIndex());

        CellStyle no06AheaderStyle = wb.createCellStyle();
        no06AheaderStyle.setFillForegroundColor(No_06A_Color.getIndex());

        CellStyle _06AheaderStyle = wb.createCellStyle();
        _06AheaderStyle.setFillForegroundColor(_06A_Color.getIndex());

        CellStyle _005headerStyle = wb.createCellStyle();
        _005headerStyle.setFillForegroundColor(_005_Color.getIndex());

        CellStyle _031headerStyle = wb.createCellStyle();
        _031headerStyle.setFillForegroundColor(_031_Color.getIndex());

        CellStyle _065headerStyle = wb.createCellStyle();
        _065headerStyle.setFillForegroundColor(_065_Color.getIndex());

        CellStyle _066headerStyle = wb.createCellStyle();
        _066headerStyle.setFillForegroundColor(_066_Color.getIndex());

        CellStyle _067headerStyle = wb.createCellStyle();
        _067headerStyle.setFillForegroundColor(_067_Color.getIndex());

        for (CellStyle cellStyle : Arrays.asList(headerStyle, no06AheaderStyle, _06AheaderStyle, _005headerStyle, _031headerStyle, _065headerStyle, _066headerStyle, _067headerStyle)) {
            cellStyle.setFont(headerFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

//        firstRow = sheet.createRow(0);
//        secondRow = sheet.createRow(1);
//        thirdRow = sheet.createRow(2);
//        totalsRow = sheet.createRow(NUM_HEADER_ROWS + rowData.size());
        List<Row> styledRows = Arrays.asList(secondRow, thirdRow, totalsRow);
        for (int i = 0; i < NUM_COLS; i++) {
            firstRow.createCell(i).setCellStyle(headerStyle);
            if (i < NO_06A_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(headerStyle);
            } else if (i == NO_06A_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(no06AheaderStyle);
            } else if (i < _005_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_06AheaderStyle);
            } else if (i < _031_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_005headerStyle);
            } else if (i < _065_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_031headerStyle);
            } else if (i < _066_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_065headerStyle);
            } else if (i < _067_COL) {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_066headerStyle);
            } else {
                for (Row row : styledRows) row.createCell(i).setCellStyle(_067headerStyle);
            }
        }
        firstRow.getCell(DATE_COL).setCellValue("Date");
        firstRow.getCell(FILE_COL).setCellValue("File");
        firstRow.getCell(NO_06A_COL).setCellValue("MSP Codes");
        secondRow.getCell(_06A_COL).setCellValue("06A");
        secondRow.getCell(_005_COL).setCellValue("005");
        secondRow.getCell(_031_COL).setCellValue("031");
        secondRow.getCell(_065_COL).setCellValue("065");
        secondRow.getCell(_066_COL).setCellValue("066");
        secondRow.getCell(_067_COL).setCellValue("067");
        secondRow.getCell(NO_06A_COL).setCellValue("No 06A");

        for (int i = _06A_COL; i < NUM_COLS; i++) {
            if (i % 3 == 0) thirdRow.getCell(i).setCellValue("Pre-Flight");
            if (i % 3 == 1) thirdRow.getCell(i).setCellValue("In-Flight");
            if (i % 3 == 2) thirdRow.getCell(i).setCellValue("Post-Flight");
        }

        totalsRow.getCell(FILE_COL).setCellValue("TOTAL:");
        totalsRow.getCell(FILE_COL).setCellStyle(headerStyle);

//        for (int i = NO_06A_COL; i < NUM_COLS; i++) {
//            char colChar = (char) ('A' + i);
//            String sumColFormula = "SUM(" + colChar + (NUM_HEADER_ROWS + 1) + ":" + colChar + (totalsRow.getRowNum()) + ")";
//            totalsRow.getCell(i).setCellFormula(sumColFormula);
//        }

        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, DATE_COL, DATE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, FILE_COL, FILE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), NO_06A_COL, NUM_COLS - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), NO_06A_COL, NO_06A_COL));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _06A_COL, _005_COL - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _005_COL, _031_COL - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _031_COL, _065_COL - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _065_COL, _066_COL - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _066_COL, _067_COL - 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), _067_COL, NUM_COLS - 1));
    }

    private void makeBorders() {
        PropertyTemplate pt = new PropertyTemplate();

        // Draw thick border around header and fill in light borders
        CellRangeAddress header = new CellRangeAddress(firstRow.getRowNum(), thirdRow.getRowNum(), DATE_COL, NUM_COLS - 1);
        pt.drawBorders(header, BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
        pt.drawBorders(header, BorderStyle.THIN, BorderExtent.INSIDE);

        CellRangeAddress dateTimeHeaderRange = new CellRangeAddress(firstRow.getRowNum(), thirdRow.getRowNum(), DATE_COL, DATE_COL);
        CellRangeAddress dateTimeFooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), DATE_COL, DATE_COL);
        CellRangeAddress dateTimeColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, DATE_COL, DATE_COL);

        CellRangeAddress fileHeaderRange = new CellRangeAddress(firstRow.getRowNum(), thirdRow.getRowNum(), FILE_COL, FILE_COL);
        CellRangeAddress fileFooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), FILE_COL, FILE_COL);
        CellRangeAddress fileColRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, FILE_COL, FILE_COL);

        CellRangeAddress no06AHeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), NO_06A_COL, NO_06A_COL);
        CellRangeAddress no06AFooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), NO_06A_COL, NO_06A_COL);
        CellRangeAddress no06AColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, NO_06A_COL, NO_06A_COL);

        CellRangeAddress _06AHeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _06A_COL, _005_COL - 1);
        CellRangeAddress _06AFooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _06A_COL, _005_COL - 1);
        CellRangeAddress _06AColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _06A_COL, _005_COL - 1);

        CellRangeAddress _005HeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _005_COL, _031_COL - 1);
        CellRangeAddress _005FooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _005_COL, _031_COL - 1);
        CellRangeAddress _005ColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _005_COL, _031_COL - 1);

        CellRangeAddress _031HeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _031_COL, _065_COL - 1);
        CellRangeAddress _031FooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _031_COL, _065_COL - 1);
        CellRangeAddress _031ColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _031_COL, _065_COL - 1);

        CellRangeAddress _065HeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _065_COL, _066_COL - 1);
        CellRangeAddress _065FooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _065_COL, _066_COL - 1);
        CellRangeAddress _065ColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _065_COL, _066_COL - 1);

        CellRangeAddress _066HeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _066_COL, _067_COL - 1);
        CellRangeAddress _066FooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _066_COL, _067_COL - 1);
        CellRangeAddress _066ColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _066_COL, _067_COL - 1);

        CellRangeAddress _067HeaderRange = new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), _067_COL, NUM_COLS - 1);
        CellRangeAddress _067FooterRange = new CellRangeAddress(totalsRow.getRowNum(), totalsRow.getRowNum(), _067_COL, NUM_COLS - 1);
        CellRangeAddress _067ColsRange = new CellRangeAddress(NUM_HEADER_ROWS, totalsRow.getRowNum() - 1, _067_COL, NUM_COLS - 1);

        for (CellRangeAddress range : Arrays.asList(
                dateTimeHeaderRange, dateTimeFooterRange, dateTimeColsRange,
                fileHeaderRange, fileFooterRange, fileColRange,
                no06AHeaderRange, no06AFooterRange, no06AColsRange,
                _06AHeaderRange, _06AFooterRange, _06AColsRange,
                _005HeaderRange, _005FooterRange, _005ColsRange,
                _031HeaderRange, _031FooterRange, _031ColsRange,
                _065HeaderRange, _065FooterRange, _065ColsRange,
                _066HeaderRange, _066FooterRange, _066ColsRange,
                _067HeaderRange, _067FooterRange, _067ColsRange
        )) {
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
            String sumColFormula = "SUM(" + colChar + (NUM_HEADER_ROWS + 1) + ":" + colChar + (totalsRow.getRowNum()) + ")";
            totalsRow.getCell(i).setCellFormula(sumColFormula);
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
}
