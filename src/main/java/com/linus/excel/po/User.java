package com.linus.excel.po;

import java.sql.Time;
import java.util.Calendar;
import java.util.Date;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import javax.validation.constraints.Size;

import org.hibernate.validator.constraints.Email;

import com.linus.excel.annotation.Header;
import com.linus.excel.enums.Gender;
import com.linus.excel.validation.UserChecker;

@UserChecker
public class User {
	private String lastName;
	private String firstName;
	private Integer age;
	private Gender gender;
	private String email;
	private Double shippingFee;
	private String free;
	private Boolean student;
	private Date birthday;
	private Time time;
	private double completed;
	private Calendar end;
	
	public User() {}
	
	@NotNull
	@Size(max=64)
	@Header(columnIndex=0)
	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	@Header(columnIndex=1)
	public String getLastName() {
		return lastName;
	}
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	@Header(columnIndex=2)
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	@Header(columnIndex=3)
	public Gender getGender() {
		return gender;
	}
	public void setGender(Gender gender) {
		this.gender = gender;
	}
	
	@Email
	@Header(columnIndex=4)
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	
	@Header(columnIndex=5)
	public Double getShippingFee() {
		return shippingFee;
	}

	public void setShippingFee(Double shippingFee) {
		this.shippingFee = shippingFee;
	}

	@NotNull
	@Pattern(regexp="yes|no|free", flags={Pattern.Flag.CASE_INSENSITIVE}, message="{validation.excel.user.free.message}")
	@Header(columnIndex=6)
	public String getFree() {
		return free;
	}
	public void setFree(String free) {
		this.free = free;
	}

	@Header(columnIndex=8)
	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	@Header(columnIndex=10)
	public Time getTime() {
		return time;
	}

	public void setTime(Time time) {
		this.time = time;
	}

	@Header(columnIndex=11)
	public double getCompleted() {
		return completed;
	}

	public void setCompleted(double completed) {
		this.completed = completed;
	}

	@Header(columnIndex=12)
	public Calendar getEnd() {
		return end;
	}

	public void setEnd(Calendar end) {
		this.end = end;
	}

	@Header(columnIndex=7)
	public Boolean getStudent() {
		return student;
	}

	public void setStudent(Boolean student) {
		this.student = student;
	}	
	
}
