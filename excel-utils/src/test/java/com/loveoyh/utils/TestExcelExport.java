package com.loveoyh.utils;

import com.loveoyh.pojo.BeanDemo;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * @Created by oyh.Jerry to 2020/04/23 08:14
 */
public class TestExcelExport {

	@Test
	public void testExport() throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InterruptedException, IOException {
		Map<String,String> map = new LinkedHashMap<String, String>();
		map.put("第一列","id");
		map.put("第二列","name");
		map.put("第三列","price");
		map.put("第四列","time");
		
		List<BeanDemo> data = new ArrayList<BeanDemo>();
		BeanDemo beanDemo = new BeanDemo();
		beanDemo.setId("1");
		beanDemo.setName("jerry");
		beanDemo.setPrice(2.0);
		beanDemo.setTime(new Date(System.currentTimeMillis()));
		data.add(beanDemo);
		
		Workbook wb = ExcelExportUtil.buildSheet(map,data);
		
		String fileName = new String("信息.xls");
		OutputStream os = new FileOutputStream(new File("G:\\Desktop\\"+fileName));
		wb.write(os);
		wb.close();
		os.close();
	}

}
