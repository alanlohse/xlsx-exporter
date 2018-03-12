package com.anlohse.porter.test;

import java.util.Date;

import com.anlohse.porter.annotation.Column;

public class Student {

	@Column(header = "Nome")
	private String name;
	
	@Column(header = "Data Nasc.")
	private Date birthDate;
	
	@Column(header = "Idade")
	private int age;
	
	@Column(header = "Média")
	private double avg;

	public Student(String name, Date birthDate, int age, double avg) {
		this.name = name;
		this.birthDate = birthDate;
		this.age = age;
		this.avg = avg;
	}

	public Student() {
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getBirthDate() {
		return birthDate;
	}

	public void setBirthDate(Date birthDate) {
		this.birthDate = birthDate;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public double getAvg() {
		return avg;
	}

	public void setAvg(double avg) {
		this.avg = avg;
	}

}
