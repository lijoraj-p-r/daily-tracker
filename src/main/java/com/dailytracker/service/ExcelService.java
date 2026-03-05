package com.dailytracker.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ExcelService {
    private static final String FILE_PREFIX = "DailyActivity_";
    private static final String[] HEADERS = {"START TIME", "CATEGORY", "ACTIVITY DESCRIPTION", "END TIME", "DURATION (MIN)"};
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");
    private static final int HEADER_ROW_INDEX = 3;

    private Path getFilePath() {
        String date = LocalDateTime.now().format(DATE_FMT);
        return Paths.get(FILE_PREFIX + date + ".xlsx");
    }

    public void logActivity(String category, String description) {
        Path path = getFilePath();
        File file = path.toFile();

        try (XSSFWorkbook workbook = (XSSFWorkbook) getWorkbook(file)) {
            XSSFSheet sheet = (XSSFSheet) getOrCreateSheet(workbook);
            int lastRowNum = sheet.getLastRowNum();
            
            if (lastRowNum > HEADER_ROW_INDEX) { 
                Row lastRow = sheet.getRow(lastRowNum);
                if (lastRow != null) {
                    Cell endTimeCell = lastRow.getCell(3);
                    if (endTimeCell == null || endTimeCell.getStringCellValue().isEmpty()) {
                        closeActivity(workbook, lastRow);
                    }
                }
            }

            int nextRow = Math.max(sheet.getLastRowNum() + 1, HEADER_ROW_INDEX + 1);
            Row row = sheet.createRow(nextRow);
            CellStyle dataStyle = createDataStyle(workbook, nextRow % 2 == 0);
            
            createStyledCell(row, 0, LocalDateTime.now().format(TIME_FMT), dataStyle);
            createStyledCell(row, 1, category.toUpperCase(), dataStyle);
            createStyledCell(row, 2, description, dataStyle);
            createStyledCell(row, 3, "", dataStyle); 
            createStyledCell(row, 4, "", dataStyle); 

            for(int i=0; i < HEADERS.length; i++) sheet.autoSizeColumn(i);

            updateCharts(workbook, sheet);
            saveWorkbook(workbook, file);
            System.out.println("Started: [" + category.toUpperCase() + "] " + description);

        } catch (IOException e) {
            System.err.println("Error: " + e.getMessage());
        }
    }

    public void stopCurrentActivity() {
        Path path = getFilePath();
        File file = path.toFile();
        if (!file.exists()) return;

        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > HEADER_ROW_INDEX) {
                Row lastRow = sheet.getRow(lastRowNum);
                Cell endTimeCell = lastRow.getCell(3);
                if (endTimeCell == null || endTimeCell.getStringCellValue().isEmpty()) {
                    closeActivity(workbook, lastRow);
                    updateCharts(workbook, sheet);
                    saveWorkbook(workbook, file);
                    System.out.println("Stopped current activity.");
                }
            }
        } catch (IOException e) {
            System.err.println("Error stopping activity: " + e.getMessage());
        }
    }

    private void closeActivity(Workbook workbook, Row row) {
        LocalDateTime now = LocalDateTime.now();
        String startTimeStr = row.getCell(0).getStringCellValue();
        LocalDateTime startTime = LocalDateTime.parse(
            LocalDateTime.now().format(DATE_FMT) + "T" + startTimeStr,
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss")
        );
        long durationMinutes = Duration.between(startTime, now).toMinutes();
        row.getCell(3).setCellValue(now.format(TIME_FMT));
        row.getCell(4).setCellValue(durationMinutes);
    }

    private void updateCharts(XSSFWorkbook workbook, XSSFSheet logSheet) {
        String analyticsSheetName = "Analytics";
        XSSFSheet analyticsSheet = workbook.getSheet(analyticsSheetName);
        if (analyticsSheet != null) {
            workbook.removeSheetAt(workbook.getSheetIndex(analyticsSheet));
        }
        analyticsSheet = workbook.createSheet(analyticsSheetName);

        // Calculate Summary
        Map<String, Long> summary = new HashMap<>();
        for (Row row : logSheet) {
            if (row.getRowNum() <= HEADER_ROW_INDEX) continue;
            Cell catCell = row.getCell(1);
            Cell durCell = row.getCell(4);
            if (catCell != null && durCell != null && durCell.getCellType() == CellType.NUMERIC) {
                summary.merge(catCell.getStringCellValue(), (long) durCell.getNumericCellValue(), Long::sum);
            }
        }

        // Write Summary Data for Chart
        int rowIdx = 0;
        Row header = analyticsSheet.createRow(rowIdx++);
        header.createCell(0).setCellValue("Category");
        header.createCell(1).setCellValue("Minutes");

        for (Map.Entry<String, Long> entry : summary.entrySet()) {
            Row row = analyticsSheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        // Create Pie Chart
        XSSFDrawing drawing = analyticsSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Time Distribution by Category");
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(analyticsSheet, 
                new CellRangeAddress(1, rowIdx - 1, 0, 0));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(analyticsSheet, 
                new CellRangeAddress(1, rowIdx - 1, 1, 1));

        XDDFPieChartData data = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
        data.setVaryColors(true);
        data.addSeries(categories, values);
        chart.plot(data);
    }

    private void createStyledCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private CellStyle createDataStyle(Workbook workbook, boolean isEven) {
        CellStyle style = workbook.createCellStyle();
        if (isEven) {
            style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private Sheet getOrCreateSheet(Workbook workbook) {
        if (workbook.getNumberOfSheets() > 0) return workbook.getSheetAt(0);
        Sheet sheet = workbook.createSheet("Daily Log");
        
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("DAILY PRODUCTIVITY REPORT");
        CellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        titleFont.setColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        Font hFont = workbook.createFont();
        hFont.setColor(IndexedColors.WHITE.getIndex());
        hFont.setBold(true);
        headerStyle.setFont(hFont);

        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
            cell.setCellStyle(headerStyle);
        }
        sheet.createFreezePane(0, 4);
        return sheet;
    }

    private Workbook getWorkbook(File file) throws IOException {
        return file.exists() ? new XSSFWorkbook(new FileInputStream(file)) : new XSSFWorkbook();
    }

    private void saveWorkbook(Workbook workbook, File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
    }

    public void printSummary() {
        Path path = getFilePath();
        File file = path.toFile();
        if (!file.exists()) { System.out.println("No logs for today."); return; }

        Map<String, Long> categoryDuration = new HashMap<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() <= HEADER_ROW_INDEX) continue;
                Cell catCell = row.getCell(1);
                Cell durCell = row.getCell(4);
                if (catCell != null && durCell != null && durCell.getCellType() == CellType.NUMERIC) {
                    categoryDuration.merge(catCell.getStringCellValue(), (long) durCell.getNumericCellValue(), Long::sum);
                }
            }
            System.out.println("\n=== Daily Summary ===");
            categoryDuration.forEach((cat, dur) -> System.out.printf("%-10s : %d mins%n", cat, dur));
            System.out.println("=====================\n");
        } catch (IOException e) {
            System.err.println("Error reading summary: " + e.getMessage());
        }
    }
}
