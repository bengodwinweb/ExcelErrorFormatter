import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class Main {

    public static final int FILE_COL = 1;
    public static List<BunoError> BunoErrors;

    static {
        BunoError errorNo06A = new BunoError("No_06A", FILE_COL + 1, FILE_COL + 1, IndexedColors.PLUM);
        BunoError error06A = new BunoError("06A", errorNo06A.getEndCol() + 1, errorNo06A.getEndCol() + 3, IndexedColors.CORAL);
        BunoError error005 = new BunoError("005", error06A.getEndCol() + 1, error06A.getEndCol() + 3, IndexedColors.LIGHT_GREEN);
        BunoError error031 = new BunoError("031", error005.getEndCol() + 1, error005.getEndCol() + 3, IndexedColors.LIGHT_CORNFLOWER_BLUE);
        BunoError error065 = new BunoError("065", error031.getEndCol() + 1, error031.getEndCol() + 3, IndexedColors.LIGHT_YELLOW);
        BunoError error066 = new BunoError("066", error065.getEndCol() + 1, error065.getEndCol() + 3, IndexedColors.LIGHT_ORANGE);
        BunoError error067 = new BunoError("067", error066.getEndCol() + 1, error066.getEndCol() + 3, IndexedColors.LIGHT_TURQUOISE);


        BunoErrors = new ArrayList<>(Arrays.asList(error06A, error005, error031, error065, error066, error067));
        BunoErrors.add(error06A);
    }

    public static void main(String[] args) {
        String inputFileFolder = "/Users/bgodwin/Local Files/Development Projects/java/ExcelProjectInputFiles/";
        String sourceName = "example2.xlsx";

        // read values from current workbook;
        try (InputStream inputStream = new FileInputStream(inputFileFolder + sourceName)) {
            Workbook wb = WorkbookFactory.create(inputStream);
//
//            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
//                Sheet sheet = wb.getSheetAt(i);
//                if (sheet.getSheetName().contains("BUNO")) {
//                    System.out.println("Removing sheet " + sheet.getSheetName());
//                    wb.removeSheetAt(i);
//                }
//            }

            System.out.println("Number of sheets: " + wb.getNumberOfSheets());

            int removeFrom = -1;
            for (int i = 1; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                if (sheet.getSheetName().contains("BUNO")) {
                    System.out.println("Remove from: " + i);
                    removeFrom = i;
                    break;
                }
            }

            if (removeFrom != -1) {
                while(wb.getNumberOfSheets() - 1 >= removeFrom) {
                    System.out.println("Removing " + wb.getSheetName(removeFrom));
                    wb.removeSheetAt(removeFrom);
                }
            }

            Sheet sheet = wb.getSheetAt(0);

            int firstRow = -1;
            for (int i = sheet.getFirstRowNum(); i < 100; i++) {
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

            Reader reader = new Reader(sheet, firstRow);
            Map<String, Buno> bunosMap = reader.getBunos();
            System.out.println("\nFound " + bunosMap.values().size() + " BUNOs");

            List<Buno> bunos = new ArrayList<>(bunosMap.values());
            for (Buno buno : bunos) {
//                System.out.println(buno);
                RowReader bunoReader = new RowReader(sheet, buno.getRows());

                List<CustomRowData> bunoData = bunoReader.getRecords().stream().sorted(Comparator.comparing(CustomRowData::getDate, Date::compareTo)).collect(Collectors.toList());
                System.out.println("Found " + bunoData.size() + " unique records for Buno " + buno.getName());

                SheetWriter bunoWriter = new SheetWriter(wb, buno.getName(), bunoData);
                bunoWriter.makeSheet();
            }

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            try (OutputStream fileOut = new FileOutputStream(inputFileFolder + sourceName)) {
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

    public void initialize() {

    }

}
