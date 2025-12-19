package com.example.timesheet.dto;

import jakarta.validation.constraints.*;

import java.util.List;

public class TimesheetRequest {



    @NotBlank(message = "Employee ID should not be Blank")

    private String employeeId;@NotNull(message = "Year should not be null")
    @Min(value = 1900, message = "Year should be after 1900")
    @Max(value = 2100, message = "Year should be before 2100")
    private int year;

    @Min(value = 1, message = "Month should be between 1 and 12")
    @Max(value = 12, message = "Month should be between 1 and 12")
    private int month;

    @NotBlank(message = "ManagerApproval should not be Blank")
    private String managerApproval;


    private List<String> holidays;
    private List<String> leaveDates;
    private List<String> remarks;


    // 1st weekoff saturday, 3rd weekoff saturday
    public static final int WEEKOFF_1ST_3RD = 1;

    // 2nd weekoff saturday, 4th weekoff saturday
    public static final int WEEKOFF_2ND_4TH = 2;

    // all saturday weekoff
    public static final int ALL_WEEKOFF = 3;

    // all saturday regular work
    public static final int EVERY_SATURDAY = 4;

    private int saturdayWeekoff;

    public String getEmployeeId() {
        return employeeId;
    }

    public void setEmployeeId(String employeeId) {
        this.employeeId = employeeId;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public int getMonth() {
        return month;
    }

    public void setMonth(int month) {
        this.month = month;
    }

    public String getManagerApproval() {
        return managerApproval;
    }

    public void setManagerApproval(String managerApproval) {
        this.managerApproval = managerApproval;
    }

    public List<String> getHolidays() {
        return holidays;
    }

    public void setHolidays(List<String> holidays) {
        this.holidays = holidays;
    }

    public List<String> getLeaveDates() {
        return leaveDates;
    }

    public void setLeaveDates(List<String> leaveDates) {
        this.leaveDates = leaveDates;
    }

    public List<String> getRemarks() {
        return remarks;
    }

    public void setRemarks(List<String> remarks) {
        this.remarks = remarks;
    }

    public int getSaturdayWeekoff() {
        return saturdayWeekoff;
    }

    public void setSaturdayWeekoff(int saturdayWeekoff) {
        this.saturdayWeekoff = saturdayWeekoff;
    }
}




// Sample request body
/*
{
  "employeeId": 1,
  "year": 2022,
  "month": 1,
  "holidays": ["2022-01-01", "2022-01-02"],
  "leaveDates": ["2022-01-15", "2022-01-16"],
  "remarks": ["on leave", "personal"],
  "saturdayWeekoff": 1
}
*/
