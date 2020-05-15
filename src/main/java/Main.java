import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;

public class Main {

    public static final int FILE_COL = 1;
    public static List<BunoError> BunoErrors;
    public static String DIRECTORY_NAME;

    static {
        int noErrCol = FILE_COL + 1;

        BunoError error06A = new BunoError("06A");
        BunoError error005 = new BunoError("005");
        BunoError error031 = new BunoError("031");
        BunoError error064 = new BunoError("064");
        BunoError error065 = new BunoError("065");
        BunoError error066 = new BunoError("066");
        BunoError error067 = new BunoError("067");
        BunoError error02A = new BunoError("02A");
        BunoError error02C = new BunoError("02C");
        BunoError error2A1 = new BunoError("2A1");
        BunoError error2A2 = new BunoError("2A2");
        BunoError error2A3 = new BunoError("2A3");

        BunoErrors = new ArrayList<>(Arrays.asList(
                error06A,
                error005,
                error031,
                error064,
                error065,
                error066,
                error067,
                error02A,
                error02C,
                error2A1,
                error2A2,
                error2A3
        ));

        BunoError errorNoCode1 = new BunoError("No " + BunoErrors.get(0).getCode(), noErrCol, noErrCol, IndexedColors.PLUM);
        BunoErrors.add(errorNoCode1);
    }

    public static void main(String[] args) {
        String sourceName;
        Scanner scan = new Scanner(System.in);

//        DIRECTORY_NAME = "/Users/bgodwin/Local Files/Development Projects/java/ExcelProjectInputFiles/";
//        sourceName = "example2.xlsx";

        System.out.println("\nEnter the absolute path of the source folder (e.x. /Users/user/documents/): ");
        DIRECTORY_NAME = scan.nextLine();
        if (DIRECTORY_NAME.charAt(DIRECTORY_NAME.length() - 1) != '/') DIRECTORY_NAME += "/";

        System.out.println("Enter the name of the source file (example.xsls): ");
        sourceName = scan.nextLine();


        // read values from current workbook;
        try (InputStream inputStream = new FileInputStream(DIRECTORY_NAME + sourceName)) {
            System.out.println("\nProcessing " + sourceName);

            Workbook wb = WorkbookFactory.create(inputStream);

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
            System.out.println("\nTotal of " + bunos.size() + " BUNOs\n");

            // iterate over the list and write a new sheet for each BUNO
            for (Buno buno : bunos) {
                // get the data from each BUNO by reading the rows in which is occurs
                RowReader bunoReader = new RowReader(sheet, buno.getRows());

                // sort the list by date of the flight
                List<CustomRowData> bunoData = bunoReader.getRecords().stream().sorted(Comparator.comparing(CustomRowData::getDate, Date::compareTo)).collect(Collectors.toList());
                System.out.println("Total of " + bunoData.size() + " flights for BUNO " + buno.getName());

                Workbook workbook = new XSSFWorkbook();

                SheetWriter bunoWriter = new SheetWriter(workbook, buno.getName(), bunoData);
                bunoWriter.makeSheet();
            }
        } catch (IOException e) {
            System.out.println("IOException while reading workbook " + e.getMessage());
            e.printStackTrace();
        }
    }

}
