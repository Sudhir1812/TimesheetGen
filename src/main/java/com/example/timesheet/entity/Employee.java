package com.example.timesheet.entity;

import com.fasterxml.jackson.annotation.JsonFormat;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.validation.constraints.Email;
import jakarta.validation.constraints.NotNull;
import jakarta.validation.constraints.Size;


@Entity
public class Employee {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @NotNull(message = "Employee name should not be empty")
    @Size(min = 1, message = "Employee name should not be empty")
    private String employeeName;
    @NotNull(message = "Employee ID should not be empty")
    @Size(min = 1, message = "Employee id should not be empty")
    private String employeeId;
    @Size(min = 1, message = "Email should not be empty")
    @NotNull(message = "Email should not be empty")
    @Email(message = "Invalid email format")
    private String email;
    @NotNull(message = "Join date should not be empty. It should be in the format DD-MM-YYYY")
    @Size(min = 1, message = "Joining date should not be empty")
    private String joiningDate;

    public Employee() {
    }

    public Employee(String employeeName, String employeeId, String email, String joiningDate) {
        this.employeeName = employeeName;
        this.employeeId = employeeId;
        this.email = email;
        this.joiningDate = joiningDate;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getJoiningDate() {
        return joiningDate;
    }

    public void setJoiningDate(String joiningDate) {
        this.joiningDate = joiningDate;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getEmployeeId() {
        return employeeId;
    }

    public void setEmployeeId(String employeeId) {
        this.employeeId = employeeId;
    }

    public String getEmployeeName() {
        return employeeName;
    }

    public void setEmployeeName(String employeeName) {
        this.employeeName = employeeName;
    }
}
