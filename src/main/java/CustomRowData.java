import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

@Getter
@Setter
@Accessors(chain = true)
public class CustomRowData {
    private String fileName;
    private Date date;
    private ArrayList<ErrorEvent> events;
//    private boolean[] eventArray;

    public CustomRowData() {
        this.events = new ArrayList<>();
//        this.eventArray = new boolean[19];
    }

    public CustomRowData(String fileName, Date date) {
        this();
        setFileName(fileName);
        setDate(date);
    }

    @Override
    public String toString() {
        return fileName + "\tErrors: " + Arrays.toString(getEventArray());
    }

    public boolean[] getEventArray() {
        boolean[] eventArray = new boolean[19];
        for (ErrorEvent e : events) {
            int i = 1 + (RowReader.EVENT_CODES.indexOf(e.getCode()) * 3);

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
