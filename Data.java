package com.WordExtractor;

public class Data {	
	
	public Data(){
		 Description = new StringBuffer();
	}
  
	private String ServiceName;
	private StringBuffer Description;	
	
	public String getServiceName() {
		return ServiceName;
	}
	public void setServiceName(String serviceName) {
		ServiceName = serviceName;
	}
	public StringBuffer getDescription() {
		return Description;
	}
	public void setDescription(StringBuffer description) {
		Description = description;
	}	
}
