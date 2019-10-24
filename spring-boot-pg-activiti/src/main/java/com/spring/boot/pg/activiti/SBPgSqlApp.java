package com.spring.boot.pg.activiti;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

import com.spring.boot.pg.activiti.service.EmployeeService;

@SpringBootApplication
public class SBPgSqlApp {

	public static void main(String[] args) {
		SpringApplication.run(SBPgSqlApp.class, args);
	}
	
	@Bean
	public CommandLineRunner init(final EmployeeService employeeService) {

		return new CommandLineRunner() {
			public void run(String... strings) throws Exception {
				employeeService.createEmployee();
				System.out.println("Employees Created");
			}
		};
	}
}
