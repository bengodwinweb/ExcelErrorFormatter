import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PropertyTemplate;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public class SheetWriter {
    public static final int NUM_HEADER_ROWS = 3;
    public static final int DATE_COL = 0;
    public static final int FILE_COL = Main.FILE_COL;
    public final int NUM_COLS;

    public static final int NUM_MONTHS = 12;
    public final int FIRST_TABLE_COL;
    public final int LAST_TABLE_COL;
    public final int FIRST_TABLE_ROW;
    public final int FIRST_BODY_ROW;
    public final int LAST_TABLE_ROW;

    private Workbook wb;
    private Sheet sheet;
    private String bunoName;
    private List<CustomRowData> rowData;

    private List<BunoError> localBunoErrors;
    private List<String> localBunoErrorCodes;

    private Row firstRow;
    private Row secondRow;
    private Row thirdRow;
    private Row footerRow;

    private Font boldFont;

    public SheetWriter(Workbook wb, String bunoName, List<CustomRowData> rowDataList) {
        this.wb = wb;
        this.bunoName = bunoName;
        DateTimeFormatter dtf = DateTimeFormat.forPattern("M-d-yy");
        this.sheet = wb.createSheet(dtf.print(new DateTime()));
        this.rowData = rowDataList;

        firstRow = sheet.createRow(0);
        secondRow = sheet.createRow(1);
        thirdRow = sheet.createRow(2);
        footerRow = sheet.createRow(NUM_HEADER_ROWS + rowDataList.size());

        boldFont = wb.createFont();
        boldFont.setBold(true);

        localBunoErrors = new ArrayList<>(Main.BunoErrors);
        System.out.println("Alias = " + (localBunoErrors == Main.BunoErrors));

        List<String> errorsInBuno = new ArrayList<>(Collections.singletonList(RowReader.EVENT_CODES.get(0)));

        // TODO - iterate over list of
//        for (errors in rowData.getErrors()) add error.code() to errorsInBuno
        System.out.println("Iterating over list of codes");
        for (CustomRowData rowData : rowDataList) {
            for (ErrorEvent errorEvent : rowData.getEvents()) {
                if (RowReader.EVENT_CODES.contains(errorEvent.getCode()) && !errorsInBuno.contains(errorEvent.getCode())) {
                    errorsInBuno.add(errorEvent.getCode());
                }
            }
        }

//        System.out.print("Errors for BUNO " + bunoName + ":\t");
//        for (String err : errorsInBuno) {
//            System.out.print(err + "\t");
//        }
//        System.out.println();

        // for BunoError in localBunoErrors if !errors.contains(bunoError) remove from localBunoErrors
        localBunoErrors.removeIf(bunoError -> !errorsInBuno.contains(bunoError.getCode()));

        int firstErrCol = 3;

        // iterate through the array and reset the start and end column
        for (int i = 0; i < localBunoErrors.size(); i++) {
            BunoError e = localBunoErrors.get(i);
            e.setStartCol(i * 3 + firstErrCol);
            e.setEndCol(e.getStartCol() + 2);
        }

        BunoError noError = Main.BunoErrors.get(Main.BunoErrors.size() - 1);
//        noError.setStartCol(firstErrCol - 1);
//        noError.setEndCol(noError.getStartCol());
        localBunoErrors.add(noError);

        localBunoErrorCodes = new ArrayList<>();
        System.out.print("Errors for BUNO " + bunoName + ":\t");
        for (BunoError e : localBunoErrors) {
            System.out.print(e.getCode() + " " + e.getStartCol() + " - " + e.getEndCol() + "\t");
            localBunoErrorCodes.add(e.getCode());
        }
        System.out.println();

        System.out.print("Errors in MAIN:\t");
        for (BunoError e : Main.BunoErrors) {
            System.out.print(e.getCode() + " " + e.getStartCol() + " - " + e.getEndCol() + "\t");
        }
        System.out.println();

        // initialize "constants" that are dependant upon the size of the errors list
        NUM_COLS = localBunoErrors.get(localBunoErrors.size() - 2).getEndCol() + 1;
        System.out.println("NUM_COLS = " + NUM_COLS);
        FIRST_TABLE_COL = NUM_COLS + 2;
        LAST_TABLE_COL = FIRST_TABLE_COL + 4;
        FIRST_TABLE_ROW = 1;
        FIRST_BODY_ROW = FIRST_TABLE_ROW + 2;
        LAST_TABLE_ROW = FIRST_TABLE_ROW + 2 + NUM_MONTHS + 1;
    }

    public void makeSheet() {
        makeHeaderAndFooter();
        makeDataRows();
        makeSumTable();
        makeFormulas();
        makeBorders();
//        deleteColumns();

        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();

        try (OutputStream fileOut = new FileOutputStream(Main.DIRECTORY_NAME + bunoName + ".xlsx")) {
            System.out.println("Writing " + bunoName + " to disk...");
            wb.write(fileOut);
        } catch (IOException e) {
            System.out.println("Exception trying workbook for BUNO " + bunoName);
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
    }

    private void makeHeaderAndFooter() {
        // style for all header and footer cells that are not for a specific error - background is PALE_BLUE
        CellStyle plainHeaderStyle = createHeaderStyle(IndexedColors.PALE_BLUE);

        // set the style for header and footer cells related to each error - get background color from the error
        for (BunoError e : localBunoErrors) e.setCellStyle(createHeaderStyle(e.getColor()));

        // create all cells in the first row and set the style to header style (there are no error specific cells in the first row)
        for (int i = 0; i < NUM_COLS; i++) firstRow.createCell(i).setCellStyle(plainHeaderStyle);

        // styling for second, third, and footer rows
        List<Row> styledRows = Arrays.asList(secondRow, thirdRow, footerRow);
        // set the first two cells to header style
        for (int i = 0; i <= FILE_COL; i++) {
            for (Row row : styledRows) row.createCell(i).setCellStyle(plainHeaderStyle);
        }
        // go through the errors and set the style for each col within the error's range to that error's style
        for (BunoError e : localBunoErrors) {
//            System.out.println("Initializing header columns " + e.getStartCol() + " - " + e.getEndCol());
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
        for (BunoError e : localBunoErrors) secondRow.getCell(e.getStartCol()).setCellValue(e.getCode());

        // set the values for the third row under the error code header - Pre-Flight then In-Flight then Post-Flight for each (except No_06A)
        for (int i = localBunoErrors.get(0).getStartCol(); i < NUM_COLS; i++) {
//            System.out.println("i = " + i);
            if (i % 3 == (localBunoErrors.get(0).getStartCol() % 3)) {
                thirdRow.getCell(i).setCellValue("Pre-Flight");
            }
            if (i % 3 == ((localBunoErrors.get(0).getStartCol() + 1) % 3)) {
                thirdRow.getCell(i).setCellValue("In-Flight");
            }
            if (i % 3 == ((localBunoErrors.get(0).getStartCol() + 2) % 3)) {
                thirdRow.getCell(i).setCellValue("Post-Flight");
            }
        }

        // create merged regions in first and footer rows
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, DATE_COL, DATE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), NUM_HEADER_ROWS - 1, FILE_COL, FILE_COL));
        sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), FILE_COL + 1, NUM_COLS - 1));
        sheet.addMergedRegion(new CellRangeAddress(footerRow.getRowNum(), footerRow.getRowNum(), DATE_COL, FILE_COL));

        // create merged regions in second row for each error over the span of its cols
        for (int i = 0; i < localBunoErrors.size() - 1; i++) {
            BunoError e = localBunoErrors.get(i);
            sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), e.getStartCol(), e.getEndCol()));
        }
        // merged region for No_06A error in its col from second to third row
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), localBunoErrors.get(localBunoErrors.size() - 1).getStartCol(), localBunoErrors.get(localBunoErrors.size() - 1).getEndCol()));
    }

    private void makeDataRows() {
        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("m/d/yy"));

        // light fill for body cols
        CellStyle lightCountCellStyle = wb.createCellStyle();
        lightCountCellStyle.setAlignment(HorizontalAlignment.CENTER);

        // dark fill for body cols
        CellStyle darkCountCellStyle = wb.createCellStyle();
        darkCountCellStyle.setAlignment(HorizontalAlignment.CENTER);
        darkCountCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        darkCountCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // add contents to body rows
        for (int i = 0; i < rowData.size(); i++) {
            Row row = sheet.createRow(i + NUM_HEADER_ROWS);
            addRowContents(row, rowData.get(i));
            row.getCell(DATE_COL).setCellStyle(dateStyle);
            row.getCell(FILE_COL + 1).setCellStyle(darkCountCellStyle); // set No_06A body columns to dark fill

            // starting with dark, alternate dark and light every three cols
            int count = 0;
            for (int j = FILE_COL + 2; j < NUM_COLS; j++) {
//                System.out.println("j = " + j);
                if ((count / 3) % 2 == 0) row.getCell(j).setCellStyle(darkCountCellStyle);
                else row.getCell(j).setCellStyle(lightCountCellStyle);
                count++;
            }
        }

        sheet.autoSizeColumn(DATE_COL);
        sheet.autoSizeColumn(FILE_COL);
    }

    // set the data from current rowData object to the current row

    private void addRowContents(Row row, CustomRowData rowData) {
        row.createCell(DATE_COL).setCellValue(rowData.getDate()); // col 0 - date
        row.createCell(FILE_COL).setCellValue(rowData.getFileName()); // col 1 - file name
//        boolean[] events = rowData.getEventArray(); // for every other col - put a 1 if events[col] == true, blank if false
        boolean[] events = makeEventArray(rowData);
//        System.out.println("events.length = " + events.length);
        for (int i = 0; i < events.length; i++) {
            if (events[i]) row.createCell(i + FILE_COL + 1).setCellValue(1);
            else row.createCell(i + FILE_COL + 1);
        }
    }

    private void makeSumTable() {
        CellStyle turquoiseHeaderStyle = createHeaderStyle(IndexedColors.LIGHT_TURQUOISE);
        CellStyle coralHeaderStyle = createHeaderStyle(localBunoErrors.get(0).getColor());

        for (int i = FIRST_TABLE_ROW; i <= LAST_TABLE_ROW; i++) {
            Row row = sheet.getRow(i);
            for (int j = FIRST_TABLE_COL; j <= LAST_TABLE_COL; j++) {
                Cell cell = row.createCell(j);
                if (i == FIRST_TABLE_ROW || i == FIRST_TABLE_ROW + 1 || i == LAST_TABLE_ROW) {
                    if (j <= FIRST_TABLE_COL + 1) cell.setCellStyle(turquoiseHeaderStyle);
                    else cell.setCellStyle(coralHeaderStyle);
                }
            }
        }

        // Set text for static cells
        sheet.getRow(FIRST_TABLE_ROW).getCell(FIRST_TABLE_COL).setCellValue("BUNO " + bunoName);
        sheet.getRow(FIRST_TABLE_ROW).getCell(FIRST_TABLE_COL + 2).setCellValue(localBunoErrors.get(0).getCode());
        sheet.getRow(FIRST_TABLE_ROW + 1).getCell(FIRST_TABLE_COL).setCellValue("Month");
        sheet.getRow(FIRST_TABLE_ROW + 1).getCell(FIRST_TABLE_COL + 1).setCellValue("Flights");
        sheet.getRow(FIRST_TABLE_ROW + 1).getCell(FIRST_TABLE_COL + 2).setCellValue("Pre-Flight");
        sheet.getRow(FIRST_TABLE_ROW + 1).getCell(FIRST_TABLE_COL + 3).setCellValue("In-Flight");
        sheet.getRow(FIRST_TABLE_ROW + 1).getCell(FIRST_TABLE_COL + 4).setCellValue("Post-Flight");
        sheet.getRow(LAST_TABLE_ROW).getCell(FIRST_TABLE_COL).setCellValue("Total:");


        // Merge cells in first header row
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), FIRST_TABLE_COL, FIRST_TABLE_COL + 1));
        sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), secondRow.getRowNum(), FIRST_TABLE_COL + 2, LAST_TABLE_COL));


        // Get start and end month
        DateTime endMonth = new DateTime().minusMonths(1).dayOfMonth().withMinimumValue().withTimeAtStartOfDay();
        DateTime startMonth = new DateTime(endMonth).minusMonths(NUM_MONTHS);

        CellStyle dateStyle = wb.createCellStyle();
        dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("mmm-yy"));

        // populate first column in body with month formatted "Jan-19"
        for (int i = 0; i <= NUM_MONTHS; i++) {
            DateTime month = new DateTime(startMonth).plusMonths(i);
            Cell cell = sheet.getRow(i + FIRST_BODY_ROW).getCell(FIRST_TABLE_COL);
            cell.setCellValue(month.toDate());
            cell.setCellStyle(dateStyle);
        }
    }

    private void makeBorders() {
        PropertyTemplate pt = new PropertyTemplate();

        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();

        // Draw thick border around outside and fill in light borders
        CellRangeAddress header = new CellRangeAddress(firstRow.getRowNum(), thirdRow.getRowNum(), DATE_COL, NUM_COLS - 1); // all of header region
        cellRangeAddresses.add(header);
        cellRangeAddresses.addAll(createRange(DATE_COL, DATE_COL)); // around date header, footer, and body cols
        cellRangeAddresses.addAll(createRange(FILE_COL, FILE_COL)); // around file header, footer, and body cols

        // borders for header, footer, and body for each error group
        for (BunoError e : localBunoErrors) cellRangeAddresses.addAll(createRange(e));

        // Table borders
        cellRangeAddresses.addAll(createTableRanges(FIRST_TABLE_ROW, LAST_TABLE_ROW, FIRST_TABLE_COL, LAST_TABLE_COL));

        for (CellRangeAddress range : cellRangeAddresses) {
            pt.drawBorders(range, BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
            pt.drawBorders(range, BorderStyle.THIN, BorderExtent.INSIDE);
        }
        pt.applyBorders(sheet);
    }

    private void makeFormulas() {
        // set the formulas for the footer row of main data
        for (int i = FILE_COL + 1; i < NUM_COLS; i++) {
            String colLetter = CellReference.convertNumToColString(i);
            String sumColFormula = "SUM(" + colLetter + (NUM_HEADER_ROWS + 1) + ":" + colLetter + (footerRow.getRowNum()) + ")"; // sum from first body row to last body row in the current column
            footerRow.getCell(i).setCellFormula(sumColFormula);
        }

        // SUMMATION TABLE
        BunoError bunoError = localBunoErrors.get(0);
        DateTimeFormatter dtf = DateTimeFormat.forPattern("YYYY,M,d"); // date format for excel formula strings

        int sourceStartRow = NUM_HEADER_ROWS + 1;
        int sourceEndRow = footerRow.getRowNum();
        String sourceDateColString = CellReference.convertNumToColString(DATE_COL);
        String sourceDateRange = sourceDateColString + sourceStartRow + ":" + sourceDateColString + sourceEndRow;

        // iterate down body of table and add formulas in each row
        for (int rowNum = FIRST_BODY_ROW; rowNum < LAST_TABLE_ROW; rowNum++) {
            DateTime month = new DateTime(sheet.getRow(rowNum).getCell(FIRST_TABLE_COL).getDateCellValue()).withTimeAtStartOfDay();
            String startDate = dtf.print(month);
            String endDate = dtf.print(month.plusMonths(1).minusMinutes(1));

            // COUNTIFS formula for total flights in a month
            String countFormula = "COUNTIFS(" + sourceDateRange + ", \">=\" & DATE(" + startDate + "), " + sourceDateRange + ", \"<=\" & DATE(" + endDate + "))";
            sheet.getRow(rowNum).getCell(FIRST_TABLE_COL + 1).setCellFormula(countFormula);

            // SUMIFS formula for
            for (int j = 0; j < 3; j++) {
                int currentCellNum = j + FIRST_TABLE_COL + 2;
                int sourceErrorCol = bunoError.getStartCol() + j;

                String sourceErrorColString = CellReference.convertNumToColString(sourceErrorCol);
                String sourceErrorRange = sourceErrorColString + sourceStartRow + ":" + sourceErrorColString + sourceEndRow;
                String sumFormula = "SUMIFS(" + sourceErrorRange + ", " + sourceDateRange + ", \">=\" & DATE(" + startDate + "), " + sourceDateRange + ", \"<=\" & DATE(" + endDate + "))";

                sheet.getRow(rowNum).getCell(currentCellNum).setCellFormula(sumFormula);
            }
        }

        // SUM formulas for footer of table
        for (int i = FIRST_TABLE_COL + 1; i <= LAST_TABLE_COL; i++) {
            String colLetter = CellReference.convertNumToColString(i);
            String sumColFormula = "SUM(" + colLetter + (FIRST_BODY_ROW + 1) + ":" + colLetter + (LAST_TABLE_ROW) + ")"; // sum from first body row to last body row in the current column
            sheet.getRow(LAST_TABLE_ROW).getCell(i).setCellFormula(sumColFormula);
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

    private List<CellRangeAddress> createTableRanges(int startRow, int endRow, int startCol, int endCol) {
        CellRangeAddress header = new CellRangeAddress(startRow, startRow + 1, startCol, startCol + 1);
        CellRangeAddress header2 = new CellRangeAddress(startRow, startRow + 1, startCol + 2, endCol);
        CellRangeAddress body = new CellRangeAddress(startRow + 2, endRow - 1, startCol, startCol + 1);
        CellRangeAddress body2 = new CellRangeAddress(startRow + 2, endRow - 1, startCol + 2, endCol);
        CellRangeAddress footer = new CellRangeAddress(endRow, endRow, startCol, startCol + 1);
        CellRangeAddress footer2 = new CellRangeAddress(endRow, endRow, startCol + 2, endCol);
        return Arrays.asList(header, header2, body, body2, footer, footer2);
    }

    private CellStyle createHeaderStyle(IndexedColors color) {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(boldFont);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(color.getIndex());
        return cellStyle;
    }

//    private List<Integer> makeErrorArrayValues(CustomRowData rowData) {
//        ArrayList<ErrorEvent> rowErrorEvents = rowData.getEvents();
//        List<Integer> eventArray = new ArrayList<>();
//        for (int i = 0; i < (localBunoErrors.size() - 1) * 3 + 1; i++) eventArray.add(0);
//
//        for (ErrorEvent e : rowErrorEvents) {
//            int i = 1 + (RowReader.EVENT_CODES.indexOf(e.getCode()) * 3);
//
//            switch (e.getMode()) {
//                case IN_FLIGHT:
//                    i++;
//                    break;
//                case POST_FLIGHT:
//                    i += 2;
//                    break;
//                default:
//                    break;
//            }
//
//            if (i >= 1 && i < eventArray.size()) eventArray.set(i, eventArray.get(i) + 1);
//        }
//        if (eventArray.get(1) + eventArray.get(2) + eventArray.get(3) == 0) eventArray.set(0, 1);
//        return eventArray;
//    }

    private boolean[] makeEventArray(CustomRowData rowData) {
        boolean[] eventArray = new boolean[(localBunoErrors.size() - 1) * 3 + 1];
        for (ErrorEvent e : rowData.getEvents()) {
            int i = 1 + (localBunoErrorCodes.indexOf(e.getCode()) * 3);

            switch (e.getMode()) {
                case UNDEFINED:
                    continue;
                case PRE_FLIGHT:
                    break;
                case IN_FLIGHT:
                    i++;
                    break;
                case POST_FLIGHT:
                    i += 2;
            }

            if (i >= 1 && i < eventArray.length) eventArray[i] = true;
        }
        if (!(eventArray[1] || eventArray[2] || eventArray[3])) eventArray[0] = true;
        return eventArray;
    }
}
