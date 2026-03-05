package com.dailytracker.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

        try (Workbook workbook = getWorkbook(file)) {
            Sheet sheet = getOrCreateSheet(workbook);
            
            // Close previous activity if exists
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

            // Create new activity row
            int nextRow = Math.max(sheet.getLastRowNum() + 1, HEADER_ROW_INDEX + 1);
            Row row = sheet.createRow(nextRow);
            
            CellStyle dataStyle = createDataStyle(workbook, nextRow % 2 == 0);
            
            createStyledCell(row, 0, LocalDateTime.now().format(TIME_FMT), dataStyle);
            createStyledCell(row, 1, category.toUpperCase(), dataStyle);
            createStyledCell(row, 2, description, dataStyle);
            createStyledCell(row, 3, "", dataStyle); // Placeholder for end time
            createStyledCell(row, 4, "", dataStyle); // Placeholder for duration

            // Auto-size columns
            for(int i=0; i < HEADERS.length; i++) sheet.autoSizeColumn(i);

            saveWorkbook(workbook, file);
            System.out.println("Started: [" + category.toUpperCase() + "] " + description);

        } catch (IOException e) {
            System.err.println("Error accessing file: " + e.getMessage());
        }
    }

    public void stopCurrentActivity() {
        Path path = getFilePath();
        File file = path.toFile();

        if (!file.exists()) {
            System.out.println("No active session found for today.");
            return;
        }

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > HEADER_ROW_INDEX) {
                Row lastRow = sheet.getRow(lastRowNum);
                Cell endTimeCell = lastRow.getCell(3);
                if (endTimeCell == null || endTimeCell.getStringCellValue().isEmpty()) {
                    closeActivity(workbook, lastRow);
                    saveWorkbook(workbook, file);
                    System.out.println("Stopped current activity.");
                } else {
                    System.out.println("No running activity to stop.");
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
        if (workbook.getNumberOfSheets() > 0) {
            return workbook.getSheetAt(0);
        } else {
            Sheet sheet = workbook.createSheet("Daily Log");
            
            // 1. Report Title
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

            // 2. Date Subtitle
            Row dateRow = sheet.createRow(1);
            Cell dateCell = dateRow.createCell(0);
            dateCell.setCellValue("Report Date: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMMM dd, yyyy")));
            CellStyle dateStyle = workbook.createCellStyle();
            Font dateFont = workbook.createFont();
            dateFont.setItalic(true);
            dateStyle.setFont(dateFont);
            dateCell.setCellStyle(dateStyle);
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 4));

            // 3. Main Table Headers
            Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.MEDIUM);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            
            Font headerFont = workbook.createFont();
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerStyle);
            }

            sheet.createFreezePane(0, 4); // Freeze Title and Headers
            return sheet;
        }
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
        if (!file.exists()) {
            System.out.println("No logs for today.");
            return;
        }

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
