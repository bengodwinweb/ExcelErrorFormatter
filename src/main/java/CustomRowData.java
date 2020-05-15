import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class CustomRowData {
    private String fileName;
    private Date date;
    private List<ErrorEvent> events;

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

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public List<ErrorEvent> getEvents() {
        return events;
    }

    public void setEvents(List<ErrorEvent> events) {
        this.events = events;
    }
}
