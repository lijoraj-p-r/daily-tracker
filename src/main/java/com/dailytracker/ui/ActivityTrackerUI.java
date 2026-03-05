package com.dailytracker.ui;

import com.dailytracker.service.CalendarService;
import com.dailytracker.service.ExcelService;
import com.dailytracker.service.ReportingService;
import java.util.Scanner;

public class ActivityTrackerUI {
    private final ExcelService excelService;
    private final ReportingService reportingService;
    private final CalendarService calendarService;
    private final Scanner scanner;

    public ActivityTrackerUI() {
        this.excelService = new ExcelService();
        this.reportingService = new ReportingService();
        this.calendarService = new CalendarService();
        this.scanner = new Scanner(System.in);
    }

    public void start() {
        System.out.println("=== Productivity Tracker 3.0 ===");
        System.out.println("Commands:");
        System.out.println("  start <CAT> <DESC>    - Begin task");
        System.out.println("  stop                  - End current task");
        System.out.println("  summary               - Current stats (CLI)");
        System.out.println("  report <WEEK/MONTH>   - Generate aggregate Excel report");
        System.out.println("  export                - Generate .ics file for Google Calendar");
        System.out.println("  exit                  - Quit");
        System.out.println("================================");

        while (true) {
            System.out.print("> ");
            String input = scanner.nextLine().trim();
            if (input.isEmpty()) continue;

            String[] parts = input.split(" ", 2);
            String command = parts[0].toLowerCase();

            switch (command) {
                case "start" -> handleStart(parts);
                case "stop" -> excelService.stopCurrentActivity();
                case "summary" -> excelService.printSummary();
                case "report" -> handleReport(parts);
                case "export" -> calendarService.exportToICS();
                case "exit" -> { System.out.println("Goodbye!"); return; }
                default -> System.out.println("Unknown command.");
            }
        }
    }

    private void handleStart(String[] parts) {
        if (parts.length < 2) { System.out.println("Usage: start <CAT> <DESC>"); return; }
        String[] args = parts[1].split(" ", 2);
        if (args.length < 2) { System.out.println("Usage: start <CAT> <DESC>"); return; }
        excelService.logActivity(args[0], args[1]);
    }

    private void handleReport(String[] parts) {
        if (parts.length < 2) { System.out.println("Usage: report <WEEK/MONTH>"); return; }
        reportingService.generateReport(parts[1].toLowerCase());
    }
}
