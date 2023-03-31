package com.bmstechpro.ensemblereports;
/* ensemble-reports
 * @created 11/10/2022
 * @author Konstantin Staykov
 */

import javafx.concurrent.Task;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReportLoader extends Task<Void> {
    private final File dataLogFile;
    private final List<Point> points = new ArrayList<>();


    public ReportLoader(File dataLogFile) {
        this.dataLogFile = dataLogFile;
    }

    private void load() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(dataLogFile))) {
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                // get points names and create Point objects
                if (row.getRowNum() == 0) {
                    row.forEach(cell -> {
                        if (cell.getColumnIndex() != 0) {
                            points.add(new Point(cell.getStringCellValue(), cell.getColumnIndex()));
                        }
                    });
                    continue;
                }

                LocalDateTime localDateTime = null;

                for (Cell cell : row) {
                    int columnIndex = cell.getColumnIndex();

                    CellType cellType = cell.getCellType();

                    switch (cellType) {

//                        case _NONE -> {
//                        }
                        case NUMERIC -> {
                            if (localDateTime != null) {
                                points.get(columnIndex - 1).addLogValue(localDateTime, cell.getNumericCellValue());
                            }
                        }
                        case STRING -> {
                            try {
                                double v = Double.parseDouble(cell.getStringCellValue());
                                points.get(columnIndex - 1).addLogValue(localDateTime, v);
                            } catch (NumberFormatException ignored) {
                            }
                        }
                        case FORMULA -> {
                            String cellValue = cell.getCellFormula();
                            localDateTime = getLocalDateTime(cellValue);
                        }
                    }
                }

            }

        }
    }

    private LocalDateTime getLocalDateTime(String str) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy h:mm:ss a");
        Pattern pattern = Pattern.compile("\"(\\d{2}/\\d{2}/\\d{4} \\d{1,2}:\\d{2}:\\d{2} [AP]M)\"");
        Matcher matcher = pattern.matcher(str);
        if (matcher.find()) {
            LocalDateTime dateTime = LocalDateTime.parse(matcher.group(1), formatter);
            int minute = dateTime.getMinute();
            int roundedMinute = minute < 15 ? 0 : minute < 30 ? 15 : minute < 45 ? 30 : 45;
            return LocalDateTime.of(dateTime.getYear(), dateTime.getMonth(), dateTime.getDayOfMonth(), dateTime.getHour(), roundedMinute);
        }


        if (str.contains("DATE") && str.contains("TIME")) {
            String[] split = str.split("\\+");
            String dateString = split[0];
            String substring = dateString.substring(dateString.indexOf("(") + 1, dateString.length() - 1);
            String timeString = split[1];
            String substring1 = timeString.substring(5, timeString.length() - 1);
            LocalDateTime dateTime = LocalDateTime.parse(substring + " " + substring1, DateTimeFormatter.ofPattern("yyyy,M,d H,m,s"));
            int minute = dateTime.getMinute();
            int roundedMinute = minute < 15 ? 0 : minute < 30 ? 15 : minute < 45 ? 30 : 45;
            return LocalDateTime.of(dateTime.getYear(), dateTime.getMonth(), dateTime.getDayOfMonth(), dateTime.getHour(), roundedMinute);
        }

        throw new IllegalArgumentException("Date Time format unknown");
    }



    private void save() throws IOException {
        String fileName = dataLogFile.getName();
        String filePath = dataLogFile.getParent() + "/" +
                fileName.replace(".xlsx", "_mod.xlsx");
        try (FileOutputStream fos = new FileOutputStream(filePath);
             Workbook workbook = new XSSFWorkbook()) {

            // Create a new sheet
            Sheet sheet = workbook.createSheet("DATA");
            // Create column headers
            Row headerRow = sheet.createRow(0);

            headerRow.createCell(0, CellType.STRING).setCellValue("Date and Time");


            for (
                    int i = 0; i < points.size(); i++) {
                Cell cell = headerRow.createCell(i + 1, CellType.STRING);
                cell.setCellValue(points.get(i).getPointName());
            }

            CellStyle style = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();

            Map<LocalDateTime, Integer> timeIndex = new HashMap<>();

            // Writing date/time range
            // Getting the data point with the maximum number of data logs
            Optional<Point> largestSet = points.stream().max(Comparator.comparingInt(Point::getDataLogsSize));
            if (largestSet.isPresent()) {
                style.setDataFormat(format.getFormat("mm/dd/yyyy h:mm"));

                Point point = largestSet.get();

                List<Point.PointRecord> dataLogs = point.getDataLogs();

                for (int i = 0; i < point.getDataLogsSize(); i++) {
                    int rowIndex = i + 1;
                    Cell cell = sheet.createRow(rowIndex)
                            .createCell(0, CellType.NUMERIC);

                    LocalDateTime dataLogDateTime = dataLogs.get(i).dateTime();

                    cell.setCellValue(dataLogDateTime);
                    cell.setCellStyle(style);
                    timeIndex.put(dataLogDateTime, rowIndex);
                }
            }

            points.forEach(dataPoint -> {
                int columnIndex = dataPoint.getColumnIndex();
                dataPoint.getDataLogs().forEach(dataLog -> {
                    Integer rowIndex = timeIndex.get(dataLog.dateTime());
                    if (rowIndex != null) {
                        Cell cell = sheet.getRow(rowIndex).createCell(columnIndex, CellType.NUMERIC);
                        cell.setCellValue(dataLog.value());
                    }

                });
            });
            sheet.setDefaultColumnWidth(15);

            sheet.autoSizeColumn(0);
            new DataLogScatterChart(sheet);
            workbook.write(fos);

        }

    }

    @Override
    protected Void call() throws Exception {
        if (!isCancelled()) {
            updateMessage("Converting file...");
            load();
            save();
            updateMessage("Report file was generated successfully");
        }

        return null;
    }
}
