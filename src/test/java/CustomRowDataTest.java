import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;

class CustomRowDataTest {

    @org.junit.jupiter.api.Test
    void getEventArray() {
        ErrorEvent inFlight06A = new ErrorEvent("06A", EVENT_MODE.IN_FLIGHT);
        ErrorEvent postFlight005 = new ErrorEvent("005", EVENT_MODE.POST_FLIGHT);
        ErrorEvent preFlight067 = new ErrorEvent("067", EVENT_MODE.PRE_FLIGHT);

        CustomRowData row = new CustomRowData();
        boolean[] boolArr;

        row.getEvents().add(inFlight06A);
        boolArr = getNewBoolArr();
        boolArr[2] = true;
        assertArrayEquals(boolArr, row.getEventArray(), "inFlight06A");

        row = new CustomRowData();
        row.getEvents().add(postFlight005);
        boolArr = getNewBoolArr();
        boolArr[0] = true; boolArr[6] = true;
        assertArrayEquals(boolArr, row.getEventArray(), "postFlight005");

        row = new CustomRowData();
        row.getEvents().add(preFlight067);
        boolArr = getNewBoolArr();
        boolArr[0] = true; boolArr[16] = true;
        assertArrayEquals(boolArr, row.getEventArray(), "preFlight067");

        row = new CustomRowData();
        row.getEvents().addAll(Arrays.asList(inFlight06A, postFlight005, preFlight067));
        boolArr = getNewBoolArr();
        boolArr[2] = true; boolArr[6] = true; boolArr[16] = true;
        assertArrayEquals(boolArr, row.getEventArray(), "all three declared events");
    }

    private boolean[] getNewBoolArr() {
        return new boolean[19];
    }
}