package com.jinu.excel.repository;

import org.springframework.data.repository.PagingAndSortingRepository;
import org.springframework.stereotype.Repository;

import com.jinu.excel.model.Employee;
@Repository
public interface EmployeeRepository  extends PagingAndSortingRepository<Employee, Long>{

}
