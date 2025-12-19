package com.example.timesheet.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class UiController {

    @GetMapping("/")
    public String home() {
        return "home"; // loads timesheet.html
    }

    @GetMapping("/timesheet")
    public String timesheetPage() {
        return "timesheet";
    }

    @GetMapping("/add-employee")
    public String addEmployeePage() {
        return "add-employee";
    }

    @GetMapping("/attendance")
    public String attendanceSheetPage() {
        return "attendance";
    }
}

