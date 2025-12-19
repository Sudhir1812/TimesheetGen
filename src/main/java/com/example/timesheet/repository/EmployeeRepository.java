package com.example.timesheet.repository;

import com.example.timesheet.entity.Employee;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.Optional;

public interface EmployeeRepository extends JpaRepository<Employee,Long> {

    Optional<Object> findByEmployeeId(String employeeId);

}
