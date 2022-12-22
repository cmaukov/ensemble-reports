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
import java.util.ArrayList;
import java.util.List;

public class ReportLoader extends Task<Void> {
    private LocalDateTime localDateTime;
    private final File dataLogFile;

    public ReportLoader(File dataLogFile) {
        this.dataLogFile = dataLogFile;
    }

    private void load(File file) throws IOException {
        String fileName = file.getName();
        String filePath = file.getParent() + "/" +
                fileName.replace(".xlsx", "_mod.xlsx");
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            Sheet sheet = workbook.getSheetAt(0);
            List<Point> points = new ArrayList<>();

            for (Row row : sheet) {
                // get points names and create Point objects
                if (row.getRowNum() == 0) {
                    row.forEach(cell -> {
                        if (cell.getColumnIndex() != 0) {
                            points.add(new Point(cell.getStringCellValue()));
                        }
                    });
                    continue;
                }

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
                            // =DATE(2022,9,4)+TIME(0,0,0)
                            if (cellValue.contains("DATE") && cellValue.contains("TIME")) {
                                String[] split = cellValue.split("\\+");
                                String dateString = split[0];
                                String substring = dateString.substring(dateString.indexOf("(") + 1, dateString.length() - 1);
                                String timeString = split[1];
                                String substring1 = timeString.substring(5, timeString.length() - 1);
                                LocalDateTime dateTime = LocalDateTime.parse(substring + " " + substring1, DateTimeFormatter.ofPattern("yyyy,M,d H,m,s"));
                                if (dateTime.getMinute() % 15 == 0) {
                                    localDateTime = dateTime;
                                }
                            }

                        }
//                        case BLANK -> {
//                        }
//                        case BOOLEAN -> {
//                        }
//                        case ERROR -> {
//                        }
                    }
                }

            }


            // Create a new sheet
            Sheet newSheet = workbook.createSheet(workbook.getSheetName(0) + "_DATA");
            // Create column headers
            Row headerRow = newSheet.createRow(0);

            headerRow.createCell(0, CellType.STRING).setCellValue("Date and Time");
            for (int i = 0; i < points.size(); i++) {
                Cell cell = headerRow.createCell(i + 1, CellType.STRING);
                cell.setCellValue(points.get(i).getPointName());
            }
            CellStyle style = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();

            for (int i = 0; i < points.size(); i++) {
                for (int j = 0; j < points.get(i).getDataLogs().size(); j++) {
                    Point.PointRecord pointRecord = points.get(i).getDataLogs().get(j);
                    if (i == 0) {
                        style.setDataFormat(format.getFormat("mm/dd/yyyy h:mm"));
                        Cell cell = newSheet.createRow(j + 1)
                                .createCell(i, CellType.NUMERIC);
                        cell.setCellValue(pointRecord.dateTime());
                        cell.setCellStyle(style);
                    }
                    newSheet.getRow(j + 1).createCell(i + 1, CellType.NUMERIC).setCellValue(pointRecord.value());
                }

            }
            newSheet.autoSizeColumn(0);
            new DataLogScatterChart(newSheet);

            try (FileOutputStream fos = new FileOutputStream(filePath)) {

                workbook.write(fos);
            }


        }
    }

    @Override
    protected Void call() throws Exception {
        if (!isCancelled()) {
            updateMessage("Converting file...");
            load(dataLogFile);
            updateMessage("Report file was generated successfully");
        }

        return null;
    }
}
