import java.util.ArrayList;
import java.util.List;

public class Buno {
    private String name;
    private List<Integer> rows;

    public Buno(String name) {
        this.name = name;
        rows = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Integer> getRows() {
        return rows;
    }

    @Override
    public String toString() {
        return "BUNO: " + name + ", rows: " + rows.size();
    }
}
