package com.example.timesheet.controller;

import com.example.timesheet.dto.AttendanceRequest;
import com.example.timesheet.service.AttendanceService;
import jakarta.validation.Valid;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.List;

@RestController
@RequestMapping("/api/attendance")
public class AttendanceController {

    private final AttendanceService service;

    public AttendanceController(AttendanceService service) {
        this.service = service;
    }

    /**
     * Send single employee attendance request. The method appends a sheet into
     * the monthly workbook (attendance-YYYY-MM.xlsx). It also returns the
     * workbook bytes so client can download immediately.
     */
    @PostMapping("/generate")
    public ResponseEntity<byte[]> generate(@RequestBody @Valid AttendanceRequest req) throws Exception {

        // ⭐ 1. Validate all incoming date lists
        validateDates(req);

        // 1. Generate or load master YEAR workbook
        byte[] file = service.generateAttendance(req);

        // 2. Extract only sheets for the employee
        byte[] file1 = service.extractSheetsForEmployee();
        //System.out.println("YEAR = " + req.year);

        // 3. File name
        String fileName = String.format("attendance-%d.xlsx",req.year);




        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"")
                .header("Access-Control-Expose-Headers", "Content-Disposition")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file1);
    }

//

    private void validateDates(AttendanceRequest req) {
        YearMonth ym = YearMonth.of(req.year, req.month);

        validateDateList(req.leaveDates, ym, "leaveDates");
        validateDateList(req.weekOffDates, ym, "weekOffDates");
        validateDateList(req.compOff, ym, "compOff");
        validateDateList(req.publicHolidays, ym, "publicHolidays");
    }

    private void validateDateList(List<String> dates, YearMonth ym, String fieldName) {
        if (dates == null) return;

        for (String d : dates) {
            LocalDate dateObj;

            try {
                dateObj = LocalDate.parse(d);
            } catch (Exception e) {
                throw new IllegalArgumentException(fieldName + " contains invalid date format: " + d);
            }

            if (dateObj.getYear() != ym.getYear() || dateObj.getMonthValue() != ym.getMonthValue()) {
                throw new IllegalArgumentException(
                        fieldName + " contains date not matching the request year/month: " + d
                );
            }
        }
    }

}

