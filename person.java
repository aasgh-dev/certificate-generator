package com.openDoc.testOpenWord;

public class person {
	private String name;
	private String gmail;
	
	public person(String n,String g) {
		this.name=g;
		this.gmail=n;
	}
	
	public void setName(String n) {
		this.name=n;
	}
	
	public String getName() {
		return this.name;
	}
	
	public String getGmail() {
		return gmail;
	}
	public void setGmail(String g) {
		this.gmail = g;
	} 
	
	@Override
	public String toString(){
		return name + "::" + gmail;
	}
}
