package Statistics;

import Statistics.Observations;

import java.util.ArrayList;

public class ParserObservations {
    private ArrayList<Observations> ObservationList;

    public ParserObservations() {
        this.ObservationList = new ArrayList<Observations>();
    }

    public void addObservation(String observation, int startLine, int endLine, String subject) {
        ObservationList.add(new Observations(observation, startLine, endLine, subject));
    }

    public ArrayList<Observations> getObservationList() {
        if (ObservationList.size() == 0) {
            ArrayList<Observations> result = new ArrayList<Observations>();
            result.add(new Observations(Observations.VBA_HAS_NOOBSERVATIONS,0,0,""));
            return result;
        }
        return ObservationList;
    }

    public int getObservationCount(String observation) {
        int result = 0;
        for (Observations obs : ObservationList) {
            if (obs.Observation.equals(observation)) {
                result++;
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

    public int countMacros() {
        int result = 0;
        for (Observations obs : ObservationList) {
            if (obs.Observation.equals(Observations.VBA_HAS_MACROS)) {
                result++;
            }
        }
        return result;
    }

    public int countCodeBlocks() {
        int result = 0;
        for (Observations obs : ObservationList) {
            switch (obs.Observation) {
                case Observations.VBA_HAS_SUBSTMT:
                case Observations.VBA_HAS_FUNCSTMT:
                    result++;
                    break;
            }
        }
        return result;
    }

    public int countCredentials() {
        int result = 0;
        for (Observations obs : ObservationList) {
            switch (obs.Observation) {
                case Observations.VBA_HAS_USERID:
                case Observations.VBA_HAS_PASSWORD:
                    result++;
                    break;
            }
        }
        return result;
    }

    public int countExternalLibRefs() {
        return getObservationCount(Observations.VBA_USES_EXTLIBS);
    }

    @Override
    public String toString() {
        return "ParserObservations{" +
                "ObservationList=" + ObservationList +
                "}";
    }
}
