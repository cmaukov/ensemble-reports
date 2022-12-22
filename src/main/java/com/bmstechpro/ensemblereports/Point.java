package com.bmstechpro.ensemblereports;
/* ensemble-reports
 * @created 11/10/2022
 * @author Konstantin Staykov
 */

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public class Point {
    private final String pointName;
    private final List<PointRecord> dataLogs = new ArrayList<>();

    public Point(String pointName) {
        this.pointName = pointName;
    }

    public void addLogValue(LocalDateTime localDateTime, double numericCellValue) {
        dataLogs.add(new PointRecord(localDateTime, numericCellValue));
    }

    public String getPointName() {
        return pointName;
    }

    public List<PointRecord> getDataLogs() {
        return dataLogs;
    }

    record PointRecord(LocalDateTime dateTime, double value) {

    }
}

