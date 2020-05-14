import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Reader {
    private Sheet sheet;
    private final int FIRST_ROW;

    public Reader(Sheet sheet, int firstRow) {
        this.sheet = sheet;
        this.FIRST_ROW = firstRow;
    }


    public List<Buno> getBunos() {
        Map<String, Buno> bunos = new HashMap<>();

        for (Row row : sheet) {
            if (row.getRowNum() < FIRST_ROW) continue;

            String bunoString = Integer.toString((int) (row.getCell(2, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK).getNumericCellValue()));

            if (bunos.containsKey(bunoString)) {
                bunos.get(bunoString).getRows().add(row.getRowNum());
            } else {
                Buno buno = new Buno(bunoString);
                buno.getRows().add(row.getRowNum());
                bunos.put(buno.getName(), buno);
            }
        }

        return new ArrayList<>(bunos.values());
    }
}
