import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.*;
import java.util.stream.Collectors;

public class RowReader {
//    public static final ArrayList<String> EVENT_CODES = new ArrayList<>(Arrays.asList("06A", "005", "031", "065", "066", "067"));
    public static final List<String> EVENT_CODES = Main.BunoErrors.stream().map(BunoError::getCode).collect(Collectors.toList());

    private Sheet sheet;
    private List<Integer> rows;

    public RowReader(Sheet sheet, List<Integer> rows) {
        this.rows = rows;
        this.sheet = sheet;
    }

    public List<CustomRowData> getRecords() {
        Map<String, CustomRowData> resultRows = new HashMap<>();

        for (Integer rowNum : rows) {
            Row row = sheet.getRow(rowNum);

            // get name of file from cell 5
            String fileName = row.getCell(5, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK).getStringCellValue();

            // get the eventCode - i
            String eventCode;
            Cell eventCell = row.getCell(6, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
            switch (eventCell.getCellType()) {
                case STRING:
                    eventCode = eventCell.getStringCellValue();
                    break;
                case NUMERIC:
                    eventCode = String.format("%03d", (int) eventCell.getNumericCellValue());
                    break;
                default:
                    eventCode = "";
                    break;
            }

            String eventModeString = row.getCell(8, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK).getStringCellValue().split(" ")[0];
            EVENT_MODE eventMode;
            switch (eventModeString) {
                case "PRE":
                    eventMode = EVENT_MODE.PRE_FLIGHT;
                    break;
                case "IN":
                    eventMode = EVENT_MODE.IN_FLIGHT;
                    break;
                case "POST":
                    eventMode = EVENT_MODE.POST_FLIGHT;
                    break;
                default:
                    eventMode = EVENT_MODE.UNDEFINED;
                    break;
            }

            Date eventDate;
            Cell dateCell = row.getCell(0, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
            eventDate = dateCell.getDateCellValue();

            ErrorEvent event = new ErrorEvent(eventCode, eventMode);
            CustomRowData resultRow;
            if (!resultRows.containsKey(fileName)) {
                resultRow = new CustomRowData(fileName, eventDate);
                resultRow.getEvents().add(event);
                resultRows.put(resultRow.getFileName(), resultRow);
            } else {
                resultRow = resultRows.get(fileName);
                resultRow.getEvents().add(event);
            }
        }

        return new ArrayList<>(resultRows.values());
    }
}
