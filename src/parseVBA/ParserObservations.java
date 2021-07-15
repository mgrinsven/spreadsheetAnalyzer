package parseVBA;

import java.util.ArrayList;

public class ParserObservations {
     private ArrayList<Observations> ObservationList;

    public ParserObservations() {
        this.ObservationList = new ArrayList<Observations>();
    }

    public void addObservation(int observation) {
        boolean exists=false;
        for (Observations obs:ObservationList) {
            if (obs.Observation == observation) {
                exists = true;
            }
        }
        if (!exists) {
            ObservationList.add(new Observations(observation, 0));
        }
    }

    public int updateObservation(int observation) {
        int result = -1;
        boolean exists = false;
        for (Observations obs:ObservationList) {
            if (obs.Observation == observation) {
                exists = true;
                obs.Count++;
                result = obs.Count;
            }
        }
        if (!exists) {
            ObservationList.add(new Observations(observation, 1));
            result = 1;
        }
        return result;
    }

    public ArrayList<Observations> getObservationList() {
        return ObservationList;
    }

    public int getObservationCount(int observation) {
        int result = -1;
        for (Observations obs : ObservationList) {
            if (obs.Observation == observation) {
                result = obs.Count;
            }
        }
        return result;
    }

    public boolean hasObservations() {
        if (ObservationList.size() > 0) {
            return true;
        }
        return false;
    }

    public boolean hasMacros() {
        boolean result = false;
        for (Observations obs : ObservationList) {
            if (obs.Observation == Observations.VBA_HAS_MACROS) {
                result = true;
            }
        }
        return result;
    }

    public boolean hasSubStmt() {
        boolean result = false;
        for (Observations obs : ObservationList) {
            if (obs.Observation == Observations.VBA_HAS_SUBSTMT) {
                result = true;
            }
        }
        return result;
    }
}
