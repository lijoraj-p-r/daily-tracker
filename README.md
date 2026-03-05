# Productivity Tracker 2.0

A professional Java 21 Command Line Interface (CLI) application for high-performance daily tracking. It automatically logs activities, tracks durations, categorizes tasks, and saves everything into styled Excel spreadsheets.

## Key Features

- **Smart Duration Tracking:** Automatically calculates minutes spent on each activity when you start a new one or stop.
- **Categorization:** Group activities into `WORK`, `STUDY`, `HEALTH`, `BREAK`, etc., for deeper analysis.
- **Styled Excel Output:** Professional spreadsheet formatting with bold, shaded headers and frozen rows.
- **Daily Summary:** Instant breakdown of time spent per category directly in the CLI.
- **Robust Architecture:** Built with clean decoupling between data models, file services, and user interface.

## Tech Stack

- **Language:** Java 21 (Records, Enhanced Switch, String Blocks)
- **Build Tool:** Maven
- **Core Library:** [Apache POI](https://poi.apache.org/) (for Excel `.xlsx` manipulation)

## Getting Started

### 1. Build the Project
```bash
cd daily-tracker
mvn package -DskipTests
```

### 2. Run the Application
```bash
java -jar target/daily-tracker-1.0-SNAPSHOT.jar
```

## Commands

| Command | Usage | Description |
| :--- | :--- | :--- |
| **`start`** | `start <CATEGORY> <description>` | Begins a new task (automatically stops the previous one). |
| **`stop`** | `stop` | Concludes the current running activity. |
| **`summary`** | `summary` | Displays total time spent per category today. |
| **`exit`** | `exit` | Closes the tracker. |

## Excel Output Structure

The app generates files named `DailyActivity_YYYY-MM-DD.xlsx` with the following columns:
1. **Start Time** (HH:mm:ss)
2. **Category** (e.g., WORK, STUDY)
3. **Description** (Details of the task)
4. **End Time** (HH:mm:ss)
5. **Duration** (Total minutes)

## Project Structure

```text
daily-tracker/
├── src/main/java/com/dailytracker/
│   ├── model/         # Data structures (Java Records)
│   ├── service/       # Excel logic and duration calculations
│   ├── ui/            # Command handling and CLI interface
│   └── App.java       # Main entry point
├── pom.xml            # Dependencies and Build config
└── README.md          # Project documentation
```

## Development and Contributions
Feel free to fork and enhance! Potential improvements:
- Data visualization in Excel charts.
- Weekly/Monthly summary generation.
- Integration with external calendar APIs.
