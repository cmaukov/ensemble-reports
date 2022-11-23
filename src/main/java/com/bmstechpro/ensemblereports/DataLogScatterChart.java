package com.bmstechpro.ensemblereports;
/* ensemble-reports
 * @created 11/12/2022
 * @author Konstantin Staykov
 */

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public final class DataLogScatterChart {
    private final XSSFSheet dataSheet;
    private final Workbook workbook;
    private XSSFSheet chartSheet;

    public DataLogScatterChart(Sheet sheet) {
        this.dataSheet = (XSSFSheet) sheet;
        this.workbook = sheet.getWorkbook();
        chartSheet = (XSSFSheet) workbook.createSheet();
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            if (i == 0) continue;
            build(i);
            if (i % 3 == 0) {
                chartSheet = (XSSFSheet) workbook.createSheet();
            }

        }


    }

    public void build(int columnIndex) {
        int anchorRow = 0;
        if (columnIndex % 3 == 1) {
            anchorRow += 17;
        }
        if (columnIndex % 3 == 0) {
            anchorRow += 34;
        }

        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, anchorRow, 20, anchorRow + 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(dataSheet.getRow(0).getCell(columnIndex).getStringCellValue());
//        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegendPosition.TOP_RIGHT);

        XDDFValueAxis bottomAxis = chart.createValueAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Date");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
//        leftAxis.setTitle("f(x)");
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);


        XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(dataSheet, new CellRangeAddress(1, dataSheet.getLastRowNum() - 1, 0, 0));
        XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(dataSheet, new CellRangeAddress(1, dataSheet.getLastRowNum() - 1, columnIndex, columnIndex));
//            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 600, 2, 2));


        XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
        XDDFScatterChartData.Series series1 = (XDDFScatterChartData.Series) data.addSeries(xs, ys1);
//        series1.setTitle("2x", null);
        series1.setSmooth(true);
//            XDDFScatterChartData.Series series2 = (XDDFScatterChartData.Series) data.addSeries(xs, ys2);
//            series2.setTitle("3x", null);
        chart.plot(data);

        solidLineSeries(data, 0, PresetColor.RED);
//        solidLineSeries(data, 1, PresetColor.TURQUOISE);


    }


    private static void solidLineSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(fill);
        XDDFChartData.Series series = data.getSeries(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setLineProperties(line);
        series.setShapeProperties(properties);
    }
}