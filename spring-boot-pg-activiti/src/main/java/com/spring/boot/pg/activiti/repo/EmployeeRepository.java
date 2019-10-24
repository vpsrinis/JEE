package com.spring.boot.pg.activiti.repo;

import org.springframework.data.jpa.repository.JpaRepository;
import com.spring.boot.pg.activiti.model.Employee;

public interface EmployeeRepository extends JpaRepository<Employee, Long> {

	public Employee findByName(String name);

}