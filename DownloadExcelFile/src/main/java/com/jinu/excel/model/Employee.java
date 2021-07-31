package com.jinu.excel.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

import lombok.Data;

@Data
@Entity
@Table(name = "EMPLOYEE")
public class Employee {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name = "emp_id" ,unique=true,nullable = false)
	private long empId;
	@Column(name = "first_name" ,unique=false,nullable = false,length = 50)
    private String firstName;
	@Column(name = "last_name" ,unique=false,length = 50)
    private String lastName;
	@Column(name = "email" ,unique=true,length = 50)
    private String email;
}
