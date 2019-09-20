package com.linus.enums;

import com.fasterxml.jackson.annotation.JsonValue;

public enum Gender implements ICustomEnum {
	MALE("Male"), FEMALE("Female");
	
	private final String name;
	
	Gender(String name) {
		this.name = name;
	}

	@JsonValue
	public String getName() {
		return name;
	}

	@Override
	public String value() {
		// TODO Auto-generated method stub
		return this.name;
	}
	
	/*@JsonCreator
	public static Gender descriptionOf(String name) {
		if (name == null || name.isEmpty()) {
			return null;
		}
		
		for (Gender g : Gender.values()) {
			if (name.equalsIgnoreCase(g.value())) {
				return g;
			}
		}
		
		return null;
	}*/
}