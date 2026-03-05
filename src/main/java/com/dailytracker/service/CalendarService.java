package com.dailytracker.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class CalendarService {
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter ICS_FMT = DateTimeFormatter.ofPattern("yyyyMMdd'T'HHmmss");

    public void exportToICS() {
        String todayDate = LocalDateTime.now().format(DATE_FMT);
        String fileName = "DailyActivity_" + todayDate + ".xlsx";
        File xlsxFile = new File(fileName);

        if (!xlsxFile.exists()) {
            System.out.println("No logs for today to export.");
            return;
        }

        StringBuilder ics = new StringBuilder();
        ics.append("BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//DailyTracker//EN\n");

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(xlsxFile))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() <= 3) continue;

                String startStr = row.getCell(0).getStringCellValue();
                String category = row.getCell(1).getStringCellValue();
                String desc = row.getCell(2).getStringCellValue();
                Cell endCell = row.getCell(3);

                if (endCell == null || endCell.getStringCellValue().isEmpty()) continue;
                String endStr = endCell.getStringCellValue();

                String dtStart = todayDate.replace("-", "") + "T" + startStr.replace(":", "");
                String dtEnd = todayDate.replace("-", "") + "T" + endStr.replace(":", "");

                ics.append("BEGIN:VEVENT\n");
                ics.append("SUMMARY:").append("[").append(category).append("] ").append(desc).append("\n");
                ics.append("DTSTART:").append(dtStart).append("\n");
                ics.append("DTEND:").append(dtEnd).append("\n");
                ics.append("END:VEVENT\n");
            }
        } catch (IOException e) {
            System.err.println("Error reading Excel for ICS: " + e.getMessage());
            return;
        }

        ics.append("END:VCALENDAR");

        String icsFileName = "DailyActivities_" + todayDate + ".ics";
        try (FileWriter writer = new FileWriter(icsFileName)) {
            writer.write(ics.toString());
            System.out.println("Exported calendar: " + icsFileName);
            System.out.println("Tip: Drag this file into Google Calendar to sync your day!");
        } catch (IOException e) {
            System.err.println("Error writing ICS file: " + e.getMessage());
        }
    }
}
