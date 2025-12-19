package com.example.timesheet.dto;


import jakarta.validation.constraints.*;

import java.util.List;

public class AttendanceRequest {
    @NotNull(message = "Employee ID should not be null")
    @NotBlank(message = "Employee ID should not be blank")
    @NotEmpty(message = "Employee ID should not be empty")
    public String employeeId;
    @NotNull(message = "Year should not be null")
    @Min(value = 1900, message = "Year should be after 1900")
    @Max(value = 2100, message = "Year should be before 2100")
    public int year;
    @NotNull(message = "Month should not be null")
    @Min(value = 1, message = "Month should be between 1 and 12")
    @Max(value = 12, message = "Month should be between 1 and 12")
    public int month; // 1-12


    public List<String> leaveDates;
    public List<String> weekOffDates;
    public List<String> compOff;
    public List<String> publicHolidays;
    public String remarks; // single remark for header (optional)

    // 1st weekoff saturday, 3rd weekoff saturday
    public static final int WEEKOFF_1ST_3RD = 1;

    // 2nd weekoff saturday, 4th weekoff saturday
    public static final int WEEKOFF_2ND_4TH = 2;

    // all saturday weekoff
    public static final int ALL_WEEKOFF = 3;

    // all saturday regular work
    public static final int EVERY_SATURDAY = 4;

    public int saturdayWeekoff;
}

