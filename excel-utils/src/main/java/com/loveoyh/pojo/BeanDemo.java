package com.loveoyh.pojo;

import java.io.Serializable;
import java.util.Date;

/**
 * 测试bean对象
 * @Created by oyh.Jerry to 2020/04/23 08:18
 */
public class BeanDemo implements Serializable {
	private String id;
	private String name;
	private Double price;
	private Date time;
	
	public String getId() {
		return id;
	}
	
	public void setId(String id) {
		this.id = id;
	}
	
	public String getName() {
		return name;
	}
	
	public void setName(String name) {
		this.name = name;
	}
	
	public Double getPrice() {
		return price;
	}
	
	public void setPrice(Double price) {
		this.price = price;
	}
	
	public Date getTime() {
		return time;
	}
	
	public void setTime(Date time) {
		this.time = time;
	}
}
