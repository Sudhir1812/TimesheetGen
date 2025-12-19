package com.example.timesheet.config;

import com.example.timesheet.entity.Employee;
import com.example.timesheet.repository.EmployeeRepository;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;

@Component
public class DataLoader implements CommandLineRunner {

    private final EmployeeRepository repo;

    public DataLoader(EmployeeRepository repo){
        this.repo=repo;
    }

    @Override
    public void run(String... args) throws Exception {



        List<Employee> employees = Arrays.asList(
                new Employee("Sudhir Kumar", "13536906", "sudhir.y.kumar@accenture.com", "21-Feb-2022"),
                new Employee("Sudip Kundu", "11038403", "sudip.d.kundu@accenture.com", "05-Oct-2020"),
                new Employee("Sourav Saha", "15590", "sourav.saha@teamcomputers.com", "02-May-2022"),
                new Employee("Anil Kumar Dash", "13536128", "anil.k.dash@accenture.com", "21-Apr-2022"),
                new Employee("Ashutosh Anand", "13536160", "ashutosh.a.anand@accenture.com", "24-Nov-2020"),
                new Employee("Rahul Kumar", "13841327", "rahul.gw@accenture.com", "24-Jun-2024"),
                new Employee("Ashish Kumar", "IKT7152", "ashishinghaniya@example.com", "05-Aug-2025"),
                new Employee("Rohit Singh", "IKT7161", "gs4293773@gmail.com", "14-Aug-2025"),
                new Employee("Niraj Kumar Thakur", "TRE-604", "nirajthakurotmppl@gmail.com", "15-Mar-2022"),
                new Employee("Rohan Pal", "TRE-1815", "rohanpal2418@gmail.com", "01-Aug-2024"),
                new Employee("Nikhil Jha", "TRE-2024", "nikhiljha321@gmail.com", "01-Aug-2024"),
                new Employee("Rishav Chakraborty", "13414729", "rishav.chakraborty@accenture.com", "16-07-2025")
        );

        repo.saveAll(employees);



    }
}
