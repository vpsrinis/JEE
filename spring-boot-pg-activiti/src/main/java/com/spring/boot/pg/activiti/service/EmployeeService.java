package com.spring.boot.pg.activiti.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.spring.boot.pg.activiti.model.Employee;
import com.spring.boot.pg.activiti.repo.EmployeeRepository;

@Service
public class EmployeeService {

	@Autowired
	private EmployeeRepository employeeRepository;

	// create the list of Employees into the database who perform the task
	public void createEmployee() {

		if (employeeRepository.findAll().size() == 0) {

			employeeRepository.save(new Employee("Vimal", "Software Enginner"));
			employeeRepository.save(new Employee("Prakash", "Technical Lead"));
			employeeRepository.save(new Employee("Subramaniam", "Test Lead"));
		}
	}
}
