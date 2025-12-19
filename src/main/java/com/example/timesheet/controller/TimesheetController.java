package com.example.timesheet.controller;

import com.example.timesheet.dto.TimesheetRequest;
import com.example.timesheet.service.TimesheetService;
import jakarta.validation.Valid;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api")

public class TimesheetController {

    private final TimesheetService service;

    public TimesheetController(TimesheetService service){
        this.service=service;
    }

    @PostMapping("/generate-timesheet")

    public ResponseEntity<byte[]> generate(@RequestBody @Valid TimesheetRequest req)throws Exception{

        if (req.getHolidays() != null && req.getRemarks() != null &&
                req.getHolidays().size() != req.getRemarks().size()) {

            throw new IllegalArgumentException("Holidays and Remarks must have same count");
        }



        byte[] file = service.generateTimesheet(req);

        // Extract only sheets for the employee
        byte[] file1 = service.extractSheetsForEmployee();


        String fileName = String.format("timesheet-%d.xlsx", req.getYear());


        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"")
                .header("Access-Control-Expose-Headers", "Content-Disposition")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file1);
    }
}
