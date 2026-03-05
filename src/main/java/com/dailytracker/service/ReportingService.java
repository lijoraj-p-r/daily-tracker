package com.dailytracker.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.stream.Stream;

public class ReportingService {
    private static final String FILE_PATTERN = "DailyActivity_*.xlsx";

    public void generateReport(String type) {
        Map<String, Long> aggregateSummary = new HashMap<>();
        try (Stream<Path> paths = Files.list(Paths.get("."))) {
            paths.filter(p -> p.getFileName().toString().matches("DailyActivity_.*\\.xlsx"))
                 .forEach(path -> aggregateData(path, aggregateSummary));

            String outputFileName = type.toUpperCase() + "_Summary_" + System.currentTimeMillis() + ".xlsx";
            saveSummaryToExcel(aggregateSummary, outputFileName, type);
            System.out.println("Generated " + type + " report: " + outputFileName);
        } catch (IOException e) {
            System.err.println("Error generating report: " + e.getMessage());
        }
    }

    private void aggregateData(Path path, Map<String, Long> aggregate) {
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() <= 3) continue; // Skip headers
                Cell catCell = row.getCell(1);
                Cell durCell = row.getCell(4);
                if (catCell != null && durCell != null && durCell.getCellType() == CellType.NUMERIC) {
                    aggregate.merge(catCell.getStringCellValue(), (long) durCell.getNumericCellValue(), Long::sum);
                }
            }
        } catch (Exception e) {
            System.err.println("Skipping file " + path + " due to error.");
        }
    }

    private void saveSummaryToExcel(Map<String, Long> summary, String fileName, String type) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(type + " Summary");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Category");
            header.createCell(1).setCellValue("Total Minutes");

            int rowIdx = 1;
            for (Map.Entry<String, Long> entry : summary.entrySet()) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(entry.getKey());
                row.createCell(1).setCellValue(entry.getValue());
            }
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);

            try (FileOutputStream fos = new FileOutputStream(fileName)) {
                workbook.write(fos);
            }
        }
    }
}
