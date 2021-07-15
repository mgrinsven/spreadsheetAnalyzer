package parseVBA;

public class Observations {
    public static final int VBA_HAS_NOBSERVATIONS = 0;
    public static final int VBA_HAS_SUBSTMT = 1;
    public static final int VBA_HAS_MACROS = 2;

    int Observation;
    int Count;

    public Observations(int observation, int count) {
        Observation = observation;
        Count = count;
    }

    public int getObservation() {
        return Observation;
    }

    public int getCount() {
        return Count;
    }
}
