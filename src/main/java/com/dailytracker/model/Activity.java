package com.dailytracker.model;

import java.time.LocalDateTime;

public record Activity(
    LocalDateTime startTime,
    LocalDateTime endTime,
    String category,
    String description,
    long durationSeconds
) {}
