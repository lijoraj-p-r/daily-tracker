package com.dailytracker.service;

import com.dailytracker.model.Activity;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ExcelService {
    private static final String FILE_PREFIX = "DailyActivity_";
    private static final String[] HEADERS = {"Start Time", "Category", "Description", "End Time", "Duration (min)"};
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

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
            if (lastRowNum > 0) { // Check if not just header
                Row lastRow = sheet.getRow(lastRowNum);
                if (lastRow != null) {
                    Cell endTimeCell = lastRow.getCell(3);
                    if (endTimeCell == null || endTimeCell.getStringCellValue().isEmpty()) {
                        closeActivity(lastRow);
                    }
                }
            }

            // Create new activity
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.createCell(0).setCellValue(LocalDateTime.now().format(TIME_FMT));
            row.createCell(1).setCellValue(category.toUpperCase());
            row.createCell(2).setCellValue(description);
            
            // Auto-size columns for readability
            for(int i=0; i<3; i++) sheet.autoSizeColumn(i);

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
            if (lastRowNum > 0) {
                Row lastRow = sheet.getRow(lastRowNum);
                Cell endTimeCell = lastRow.getCell(3);
                if (endTimeCell == null || endTimeCell.getStringCellValue().isEmpty()) {
                    closeActivity(lastRow);
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

    private void closeActivity(Row row) {
        LocalDateTime now = LocalDateTime.now();
        String startTimeStr = row.getCell(0).getStringCellValue();
        
        // Parse start time (assuming today's date for simplicity in duration calc)
        LocalDateTime startTime = LocalDateTime.parse(
            LocalDateTime.now().format(DATE_FMT) + "T" + startTimeStr,
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss")
        );

        long durationMinutes = Duration.between(startTime, now).toMinutes();
        
        row.createCell(3).setCellValue(now.format(TIME_FMT));
        row.createCell(4).setCellValue(durationMinutes);
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
                if (row.getRowNum() == 0) continue; // Skip header

                Cell categoryCell = row.getCell(1);
                Cell durationCell = row.getCell(4);

                if (categoryCell != null && durationCell != null && durationCell.getCellType() == CellType.NUMERIC) {
                    String category = categoryCell.getStringCellValue();
                    long duration = (long) durationCell.getNumericCellValue();
                    categoryDuration.merge(category, duration, Long::sum);
                }
            }

            System.out.println("\n=== Daily Summary ===");
            if (categoryDuration.isEmpty()) {
                System.out.println("No completed activities yet.");
            } else {
                categoryDuration.forEach((cat, dur) -> 
                    System.out.printf("%-10s : %d mins%n", cat, dur));
            }
            System.out.println("=====================\n");

        } catch (IOException e) {
            System.err.println("Error reading summary: " + e.getMessage());
        }
    }

    private Workbook getWorkbook(File file) throws IOException {
        if (file.exists()) {
            return new XSSFWorkbook(new FileInputStream(file));
        } else {
            Workbook workbook = new XSSFWorkbook();
            return workbook;
        }
    }

    private Sheet getOrCreateSheet(Workbook workbook) {
        if (workbook.getNumberOfSheets() > 0) {
            return workbook.getSheetAt(0);
        } else {
            Sheet sheet = workbook.createSheet("Daily Log");
            Row header = sheet.createRow(0);
            
            // Create a professional header style
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            
            Font font = workbook.createFont();
            font.setBold(true);
            headerStyle.setFont(font);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = header.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerStyle);
            }
            sheet.createFreezePane(0, 1); // Freeze the header row
            return sheet;
        }
    }

    private void saveWorkbook(Workbook workbook, File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
    }
}
