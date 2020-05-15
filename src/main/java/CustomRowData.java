import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Getter
@Setter
@Accessors(chain = true)
public class CustomRowData {
    private String fileName;
    private Date date;
    private ArrayList<ErrorEvent> events;

    public CustomRowData() {
        this.events = new ArrayList<>();
    }

    public CustomRowData(String fileName, Date date) {
        this();
        setFileName(fileName);
        setDate(date);
    }

    @Override
    public String toString() {
        return fileName + "\tErrors size: " + events.size();
    }

//    public boolean[] getEventArray() {
//        boolean[] eventArray = new boolean[(RowReader.EVENT_CODES.size() - 1) * 3 + 1];
//        for (ErrorEvent e : events) {
//            int i = 1 + (RowReader.EVENT_CODES.indexOf(e.getCode()) * 3);
//
//            switch (e.getMode()) {
//                case UNDEFINED:
//                    continue;
//                case PRE_FLIGHT:
//                    break;
//                case IN_FLIGHT:
//                    i++;
//                    break;
//                case POST_FLIGHT:
//                    i += 2;
//            }
//
//            if (i >= 1 && i < eventArray.length) eventArray[i] = true;
//        }
//        if (!(eventArray[1] || eventArray[2] || eventArray[3])) eventArray[0] = true;
//        return eventArray;
//    }

    public List<Integer> getEventArrayInts() {
        List<Integer> eventArray = new ArrayList<>();
        for (int i = 0; i < (Main.BunoErrors.size() - 1) * 3 + 1; i++) eventArray.add(0);

        for (ErrorEvent e : events) {
            int i = 1 + (RowReader.EVENT_CODES.indexOf(e.getCode()) * 3);

            switch (e.getMode()) {
                case IN_FLIGHT:
                    i++;
                    break;
                case POST_FLIGHT:
                    i += 2;
                    break;
                default:
                    break;
            }

            if (i >= 1 && i < eventArray.size()) eventArray.set(i, eventArray.get(i) + 1);
        }
        if (eventArray.get(1) + eventArray.get(2) + eventArray.get(3) == 0) eventArray.set(0, 1);
        return eventArray;
    }

    public ArrayList<ErrorEvent> getEvents() {
        return events;
    }
}
