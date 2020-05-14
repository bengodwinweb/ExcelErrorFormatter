import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class Main {

    public static final int FILE_COL = 1;
    public static List<BunoError> BunoErrors;

    static {
        BunoError error06A = new BunoError("06A", FILE_COL + 2, FILE_COL + 2 + 2, IndexedColors.CORAL);
        BunoError error005 = new BunoError("005", error06A.getEndCol() + 1, error06A.getEndCol() + 3, IndexedColors.LIGHT_GREEN);
        BunoError error031 = new BunoError("031", error005.getEndCol() + 1, error005.getEndCol() + 3, IndexedColors.LIGHT_CORNFLOWER_BLUE);
        BunoError error064 = new BunoError("064", error031.getEndCol() + 1, error031.getEndCol() + 3, IndexedColors.PLUM);
        BunoError error065 = new BunoError("065", error064.getEndCol() + 1, error064.getEndCol() + 3, IndexedColors.LIGHT_YELLOW);
//        BunoError error065 = new BunoError("065", error031.getEndCol() + 1, error031.getEndCol() + 3, IndexedColors.LIGHT_YELLOW);
        BunoError error066 = new BunoError("066", error065.getEndCol() + 1, error065.getEndCol() + 3, IndexedColors.LIGHT_ORANGE);
        BunoError error067 = new BunoError("067", error066.getEndCol() + 1, error066.getEndCol() + 3, IndexedColors.LIGHT_TURQUOISE);
        BunoError error02A = new BunoError("02A", error067.getEndCol() + 1, error067.getEndCol() + 3, IndexedColors.LIGHT_GREEN);
        BunoError error02C = new BunoError("02C", error02A.getEndCol() + 1, error02A.getEndCol() + 3, IndexedColors.LIGHT_CORNFLOWER_BLUE);
        BunoError error2A1 = new BunoError("2A1", error02C.getEndCol() + 1, error02C.getEndCol() + 3, IndexedColors.LIGHT_YELLOW);
        BunoError error2A2 = new BunoError("2A2", error2A1.getEndCol() + 1, error2A1.getEndCol() + 3, IndexedColors.LIGHT_ORANGE);
        BunoError error2A3 = new BunoError("2A3", error2A2.getEndCol() + 1, error2A2.getEndCol() + 3, IndexedColors.LIGHT_TURQUOISE);


        BunoErrors = new ArrayList<>(Arrays.asList(error06A, error005, error031, error064, error065, error066, error067, error02A, error02C, error2A1, error2A2, error2A3));
//        BunoErrors = new ArrayList<>(Arrays.asList(error06A, error005, error031, error065, error066, error067));

        BunoError errorNoCode1 = new BunoError("No " + BunoErrors.get(0).getCode(), FILE_COL + 1, FILE_COL + 1, IndexedColors.PLUM);
        BunoErrors.add(errorNoCode1);
    }

    public static void main(String[] args) {
        String inputFileFolder = "/Users/bgodwin/Local Files/Development Projects/java/ExcelProjectInputFiles/";
        String sourceName = "example2.xlsx";

        // read values from current workbook;
        try (InputStream inputStream = new FileInputStream(inputFileFolder + sourceName)) {
            Workbook wb = WorkbookFactory.create(inputStream);

            // find the first file that contains BUNO.
            // If one exists, keep removing the sheet at that index until done
            int removeFrom = -1;
            for (int i = 1; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                if (sheet.getSheetName().contains("BUNO")) {
                    removeFrom = i;
                    break;
                }
            }
            if (removeFrom != -1) while(wb.getNumberOfSheets() - 1 >= removeFrom) wb.removeSheetAt(removeFrom);

            // get the first sheet in the workbook
            Sheet sheet = wb.getSheetAt(0);

            // find the first row of data
            // look for the row where the value of cell[0] is "MU_DATE", data begins on the following row
            int firstRow = -1;
            for (int i = sheet.getFirstRowNum(); i < 50; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell firstCell = row.getCell(0, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if (firstCell != null && firstCell.getCellType() == CellType.STRING) {
                        String cellVal = firstCell.getStringCellValue();
                        if (cellVal.equals("MU_DATE")) {
                            firstRow = row.getRowNum() + 1;
                            break;
                        }
                    }
                }
            }

            // read the sheet to find and create an array of the BUNOs in the sheet
            Reader reader = new Reader(sheet, firstRow);
            List<Buno> bunos = reader.getBunos();
            System.out.println("\nFound " + bunos.size() + " BUNOs\n");

            // iterate over the list and write a new sheet for each BUNO
            for (Buno buno : bunos) {
                // get the data from each BUNO by reading the rows in which is occurs
                RowReader bunoReader = new RowReader(sheet, buno.getRows());

                // sort the list by date of the flight
                List<CustomRowData> bunoData = bunoReader.getRecords().stream().sorted(Comparator.comparing(CustomRowData::getDate, Date::compareTo)).collect(Collectors.toList());
                System.out.println("Found " + bunoData.size() + " flights for BUNO " + buno.getName());

                // make the sheet
                SheetWriter bunoWriter = new SheetWriter(wb, buno.getName(), bunoData);
                bunoWriter.makeSheet();
            }

            // evaluate all formulas
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            // write to the file
            try (OutputStream fileOut = new FileOutputStream(inputFileFolder + sourceName)) {
                System.out.println("\nWriting to disk...");
                wb.write(fileOut);
            } catch (IOException e) {
                System.out.println("IOException while writing workbook " + e.getMessage());
                e.printStackTrace();
            }

        } catch (IOException e) {
            System.out.println("IOException while reading workbook " + e.getMessage());
            e.printStackTrace();
        }
    }

}
