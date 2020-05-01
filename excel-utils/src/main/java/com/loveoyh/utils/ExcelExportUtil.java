package com.loveoyh.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Excel构建工具类
 * 注：
 * 可以在map的value直接传入Number类型
 * 暴露时间格式属性，供修改
 *
 * bean的属性类型支持Double,Integer,Date,String
 * 暂无异常处理
 *
 * @Created by oyh.Jerry to 2020/04/22 17:53
 */
public class ExcelExportUtil {
	public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
	//格式化时间的格式
	public static String pattern = DEFAULT_DATE_PATTERN;
	
	
	/**
	 * 创建工作簿
	 *
	 * @param mapping 列名和需要填充数据对象(属性或者数字)的对应关系，使用的是有序集合LinkedHashMap，put顺序对应列顺序
	 * @param data 数据
	 * @return Workbook对象
	 * @throws NoSuchMethodException
	 * @throws InvocationTargetException
	 * @throws IllegalAccessException
	 */
	public static Workbook buildSheet(Map<String, Object> mapping, List<?> data){
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("sheet1");
		//构造头行
		Row headerRow = sheet.createRow(0);
		
		int index = 0;
		for (String key : mapping.keySet()) {
			Cell headerCell = headerRow.createCell(index++);
			headerCell.setCellValue(key);
		}
		//构建数据行
		for (int i = 0; i < data.size(); i++) {
			//创建一行
			Row dataRow = sheet.createRow(i + 1);
			
			int indexCell = 0;
			for (Map.Entry<String, Object> entry : mapping.entrySet()) {
				//根据Map的value值判断
				Object obj = entry.getValue();
				if(obj instanceof Number){
					Number number = (Number) obj;
					fillDataCell(dataRow, indexCell++, number);
					continue;
				}
				
				//首字母大写并添加get前缀
				String prototype = "get" + firstUpperCase((String) entry.getValue());
				//获取方法
				Object t = data.get(i);
				Method method = null;
				try {
					method = t.getClass().getMethod(prototype);
					Object value = method.invoke(t);
					//填充值到一行单元格
					fillDataCell(dataRow, indexCell++, value);
					continue;
				} catch (NoSuchMethodException e) {
					System.err.println("调用"+prototype+"方法异常！");
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					System.err.println("无权访问该方法"+method);
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				//填充空
				fillDataCell(dataRow, indexCell++, null);
			}
		}
		return wb;
	}
	
	/**
	 * 填充每行的数据
	 *
	 * @param dataRow   行对象
	 * @param cellIndex 单元格索引
	 * @param value     将被填入的值对象
	 */
	private static void fillDataCell(Row dataRow, int cellIndex, Object value) {
		//创建一个单元格
		Cell cell = dataRow.createCell(cellIndex);
		//填充数据
		if (null != value) {
			if (value instanceof Double) {
				cell.setCellValue((Double) value);
			} else if (value instanceof Integer) {
				cell.setCellValue((Integer) value);
			} else if (value instanceof Date) {
				Date date = (Date) value;
				DateFormat df = new SimpleDateFormat(pattern);
				cell.setCellValue(df.format(date));
			} else {
				cell.setCellValue(String.valueOf(value));
			}
		}
	}
	
	/**
	 * 首字母大写转换
	 *
	 * @param str 将要转换的字符串
	 * @return 首字母大写的字符串
	 */
	public static String firstUpperCase(String str) {
		char[] ch = str.toCharArray();
		if (ch[0] >= 'a' && ch[0] <= 'z') {
			ch[0] = (char) (ch[0] - 32);
		}
		return new String(ch);
	}
}
