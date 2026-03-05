package com.dailytracker.ui;

import com.dailytracker.service.ExcelService;
import java.util.Scanner;

public class ActivityTrackerUI {
    private final ExcelService excelService;
    private final Scanner scanner;

    public ActivityTrackerUI() {
        this.excelService = new ExcelService();
        this.scanner = new Scanner(System.in);
    }

    public void start() {
        System.out.println("=== Productivity Tracker 2.0 ===");
        System.out.println("Commands:");
        System.out.println("  start <category> <description>  (e.g., 'start WORK Coding')");
        System.out.println("  stop                            (Stop current activity)");
        System.out.println("  summary                         (Show today's stats)");
        System.out.println("  exit                            (Quit app)");
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
                case "exit" -> {
                    System.out.println("Goodbye!");
                    return;
                }
                default -> System.out.println("Unknown command. Try 'start', 'stop', 'summary', or 'exit'.");
            }
        }
    }

    private void handleStart(String[] parts) {
        if (parts.length < 2) {
            System.out.println("Usage: start <category> <description>");
            return;
        }
        
        String[] args = parts[1].split(" ", 2);
        if (args.length < 2) {
            System.out.println("Usage: start <category> <description>");
            return;
        }

        String category = args[0];
        String description = args[1];
        excelService.logActivity(category, description);
    }
}
