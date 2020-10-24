package com.info.reader.model;

import java.time.LocalDate;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Employee {

	private String name;
	private String email;
	private LocalDate dateOfJoining;
	private int salary;
	private String department;


}
