package com.bmstechpro.ensemblereports;
/* ensemble-reports
 * @created 11/10/2022
 * @author Konstantin Staykov
 */

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public class Point implements Comparable<Point> {
    private final String pointName;
    private final int columnIndex;
    private final List<PointRecord> dataLogs = new ArrayList<>();

    public Point(String pointName, int columnIndex) {
        this.pointName = pointName;
        this.columnIndex = columnIndex;
    }

    public void addLogValue(LocalDateTime localDateTime, double numericCellValue) {
        dataLogs.add(new PointRecord(localDateTime, numericCellValue));
    }

    public String getPointName() {
        return pointName;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public List<PointRecord> getDataLogs() {
        return dataLogs;
    }
    public int getDataLogsSize(){
        return dataLogs.size();
    }

    @Override
    public int compareTo(Point o) {
        return Integer.compare(o.getDataLogsSize(),getDataLogsSize());
    }

    record PointRecord(LocalDateTime dateTime, double value) {

    }
}

