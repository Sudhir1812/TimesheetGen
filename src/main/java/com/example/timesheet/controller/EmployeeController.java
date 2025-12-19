package com.example.timesheet.controller;



import com.example.timesheet.entity.Employee;
import com.example.timesheet.repository.EmployeeRepository;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/api/employees")
//@CrossOrigin("*")
public class EmployeeController {

    private final EmployeeRepository repo;

    public EmployeeController(EmployeeRepository repo) {
        this.repo = repo;
    }

    @PostMapping
    public ResponseEntity<String> saveEmployee(@RequestBody @Valid Employee emp) {

        boolean existingEmployee = repo.findByEmployeeId(emp.getEmployeeId()).isPresent();

        if (existingEmployee){
            return ResponseEntity.badRequest().body("Employee already exists");

        }

        LocalDate currentDate = LocalDate.now();
        String dateFormat = "dd-MMM-yyyy";
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(dateFormat);
        try {
            currentDate = LocalDate.parse(currentDate.format(formatter), formatter);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body("Invalid format for current date");
        }

        LocalDate joiningDate;
        try {
            joiningDate = LocalDate.parse(emp.getJoiningDate(), formatter);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body("Invalid format for joining date");
        }
        if (joiningDate.isAfter(currentDate)) {
            return ResponseEntity.badRequest().body("Joining date cannot be a future date");
        }
        // SAVE TO DB OR FILE
        repo.save(emp);
        return ResponseEntity.ok("Employee saved");
    }

}

