package com.loveoyh.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Excel构建工具类
 * 注：
 * bean的属性类型支持Double,Integer,Date,String
 * 暂无异常处理
 *
 * @Created by oyh.Jerry to 2020/04/22 17:53
 */
public class ExcelExportUtil {
	public static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
	/** 格式化时间的格式 */
	public String pattern = DEFAULT_DATE_PATTERN;
	
	/** 当前的工作簿对象 */
	private Workbook wb;
	
	/** 当前正在操作的sheet */
	private String currentSheetName;
	
	/** 保存开始行和列,默认0 从0开始 */
	private Integer startRow = 0;
	private Short startCol = 0;
	/** 通用单元格样式 */
	private CellStyle cellStyle;
	
	/** 保存表格最大行列 */
	private Integer lastRow = startRow;
	private Short lastCol = startCol;
	/** 当前在第几行 */
	private Integer index = lastRow;
	
	/** 多行表格标题的分割符用于切割 */
	private String headSeparator;
	
	public ExcelExportUtil() {
		this(new HSSFWorkbook());
	}
	
	public ExcelExportUtil(Workbook wb){
		this(wb,DEFAULT_DATE_PATTERN);
	}
	
	public ExcelExportUtil(Workbook wb, String pattern) {
		this.wb = wb;
		this.pattern = pattern;
		this.headSeparator = "\\s";
		// 创建单元格样式
		this.cellStyle = wb.createCellStyle();
		// 设置单元格水平方向对其方式
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		// 设置单元格垂直方向对其方式
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
		cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
		cellStyle.setBorderTop(BorderStyle.THIN);//上边框
		cellStyle.setBorderRight(BorderStyle.THIN);//右边框
	}
	
	public Workbook buildSheet(Map<String, String> mapping, List<?> data){
		return buildSheet(null,mapping,data);
	}
	
	/**
	 * 创建工作簿
	 *
	 * 如果map的value是属性值,则会通过属性去调用对应的属性值填充格子,如果根据属性找不到对应的get方法,则会直接写入该value值,对应的
	 * 		data是对应类型的数据集合
	 * 如果map的value是字符串类型可以转Integer,则对应的data应该为字符串数组集合,此时map对应的value是data集合中的String数组的索引
	 * 		来获取该值,如果下标越界,则直接写入map对应的value值
	 * 其他的则直接写入map的value值
	 * @param name 创建Sheet表格的名字
	 * @param mapping 列名和属性名的对应关系，使用的是有序集合LinkedHashMap，put顺序对应列顺序
	 * @param data    数据
	 * @return Workbook对象
	 */
	public Workbook buildSheet(String name,Map<String, String> mapping, List<?> data){
		//构建sheet
		Sheet sheet = null;
		if(name == null){
			sheet = wb.createSheet();
		}else{
			sheet = wb.getSheet(name);
		}
		if(sheet == null){
			sheet = wb.createSheet(name);
		}
		//每次操作将当前操作的sheet表名保存
		this.currentSheetName = sheet.getSheetName();
		
		//构造表格head行
		buildHeader(sheet, mapping);
		
		//构建数据行
		for (Object obj : data) {
			//创建一行
			Row dataRow = sheet.getRow(index);
			if(dataRow == null){
				dataRow = sheet.createRow(index++);
			}
			
			int indexCell = this.startCol;
			for (Map.Entry<String, String> entry : mapping.entrySet()) {
				Object value = null;
				//首字母大写并添加get前缀
				String prototype = "get" + firstUpperCase(entry.getValue());
				//获取方法
				Method method = null;
				try {
					method = obj.getClass().getMethod(prototype);
					value = method.invoke(obj);
				} catch (NoSuchMethodException e) {
					// 如果map的值为数字,尝试转换为int,若转换失败则为-1,则直接将map的值写进cell中
					// 若转换成功,则通过index获取数据中的对应下标的值,如果下标越界,则直接将map的值写进cell中
					int index = Integer.getInteger(entry.getValue());
					try {
						value = index == -1 ? entry.getValue() : data.get(index);
					} catch (IndexOutOfBoundsException ee){
						value = entry.getValue();
					}
				} catch (IllegalAccessException e) {
					System.err.println("无权访问该方法"+method);
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				}
				//填充值到一行单元格
				fillDataCell(dataRow, indexCell++, value);
			}
			
			//记录列的最大值
			if(dataRow.getLastCellNum()>this.lastCol){
				this.lastCol = dataRow.getLastCellNum();
			}
		}
		
		//记录末行
		this.lastRow = sheet.getLastRowNum();
		
		//表格调整
		fixSheet(sheet);
		
		return wb;
	}
	
	/**
	 * 调整表格样式
	 * @param sheet
	 */
	private void fixSheet(Sheet sheet) {
		for(int i=0; i<this.lastCol; i++){
			sheet.autoSizeColumn(i);
			sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 12 / 10);
		}
	}
	
	/**
	 * 构造head行
	 * @param sheet
	 * @param mapping
	 */
	private void buildHeader(Sheet sheet, Map<String, String> mapping) {
		Map<String, String[]> headers = new HashMap<String, String[]>();
		//记录有多少行head头
		int headRows = Integer.MIN_VALUE;
		
		for (String key : mapping.keySet()) {
			String[] heads = key.split("\\s");
			if(headRows<heads.length){
				headRows = heads.length;
			}
			headers.put(key,heads);
		}
		
		for(int i=0;i<headRows;i++){
			//构造头行
			Row headerRow = sheet.getRow(index);
			if(headerRow == null){
				headerRow = sheet.createRow(index++);
			}
			
			//设置起始列
			int col = this.startCol;
			
			for (String key : mapping.keySet()) {
				Cell headerCell = headerRow.createCell(col);
				String[] heads = headers.get(key);
				headerCell.setCellValue(heads[i>=heads.length ? heads.length-1 : i]);
				headerCell.setCellStyle(this.cellStyle);
				col++;
			}
			
			//记录末列
			if(headerRow.getLastCellNum()>this.lastCol){
				this.lastCol = headerRow.getLastCellNum();
			}
		}
		
		//记录末行
		this.lastRow = sheet.getLastRowNum();
	}
	
	/**
	 * 合并指定范围行列的相同行
	 * @param name 表格的名字
	 * @param startRow 开始行,0-
	 * @param endRow 结束行
	 * @param startCol 开始列 0-
	 * @param endCol 结束列
	 * @return
	 */
	public Workbook mergeRows(String name,Integer startRow,Integer endRow,Integer startCol,Integer endCol){
		startRow = startRow==null ? this.startRow : startRow;
		endRow = endRow==null ? this.lastRow : endRow;
		startCol = startCol==null ? this.startCol : startCol;
		endCol = endCol==null ? this.lastCol : endCol;
		
		Sheet sheet = this.wb.getSheet(name);
		
		startRow = this.startRow<=startRow ? startRow : this.startRow;
		if(startRow>=this.lastRow) {
			return this.wb;
		}
		endRow = endRow>this.lastRow ? this.lastRow : endRow;
		startCol = this.startCol<=startCol ? startCol : this.startCol;
		if(startCol>this.lastCol){
			return this.wb;
		}
		endCol = endCol>this.lastCol ? this.lastCol : endCol;
		
		for(int i=startCol; i <= endCol; i++){
			String old = UUID.randomUUID().toString();
			int start = startRow;
			int row = start;
			
			while(row<=endRow){
				Row rowTemp = sheet.getRow(row);
				//处理多个table之间的空行
				if(rowTemp == null){
					if(row - start > 1){
						CellRangeAddress cra =new CellRangeAddress(start,row-1,i,i);
						sheet.addMergedRegion(cra);
					}
					start = row;
					//保证老的值在遇到存在行时得到赋值存在cell的值
					old = UUID.randomUUID().toString();
					
					row++;
					continue;
				}
				
				Cell cell = rowTemp.getCell(i);
				//处理不存在的cell
				if(cell == null){
					if(row - start > 1){
						CellRangeAddress cra =new CellRangeAddress(start,row-1,i,i);
						sheet.addMergedRegion(cra);
					}
					start = row;
					//保证老的值在遇到存在行时得到赋值存在cell的值
					old = UUID.randomUUID().toString();
					
					row++;
					continue;
				}
				
				String value = cell.getStringCellValue();
				if(!old.equals(value)){
					if(row - start > 1){
						CellRangeAddress cra =new CellRangeAddress(start,row-1,i,i);
						sheet.addMergedRegion(cra);
					}
					old = value;
					start = row;
				}
				if(row == endRow){
					if(row - start >= 1){
						CellRangeAddress cra =new CellRangeAddress(start,row,i,i);
						sheet.addMergedRegion(cra);
					}
				}
				row++;
			}
		}
		return wb;
	}
	public Workbook mergeRows(String name){
		return mergeRows(name,null,null,null,null);
	}
	
	/**
	 * 合并指定范围行列的相同列
	 * @param name 表格的名字
	 * @param startRow 开始行,0-
	 * @param endRow 结束行
	 * @param startCol 开始列 0-
	 * @param endCol 结束列
	 * @return
	 */
	public Workbook mergeCols(String name,Integer startRow,Integer endRow,Integer startCol,Integer endCol){
		startRow = startRow==null ? this.startRow : startRow;
		endRow = endRow==null ? this.lastRow : endRow;
		startCol = startCol==null ? this.startCol : startCol;
		endCol = endCol==null ? this.lastCol : endCol;
		
		Sheet sheet = this.wb.getSheet(name);
		
		startRow = this.startRow<=startRow ? startRow : this.startRow;
		if(startRow>this.lastRow) {
			return this.wb;
		}
		endRow = endRow>this.lastRow ? this.lastRow : endRow;
		
		startCol = this.startCol<=startCol ? startCol : this.startCol;
		if(this.lastCol <= startCol) {
			return this.wb;
		}
		endCol = endCol>this.lastCol ? sheet.getRow(startRow).getLastCellNum()-1 : endCol;
		
		for(int i=startRow; i <= endRow;i++){
			String old = UUID.randomUUID().toString();
			int start = startCol;
			int col = start;
			
			while(col<=endCol){
				Row rowTemp = sheet.getRow(i);
				//处理多个table之间的空列
				if(rowTemp == null){
					if(col - start > 1){
						CellRangeAddress cra =new CellRangeAddress(i,i,start,col-1);
						sheet.addMergedRegion(cra);
					}
					start = col;
					//保证老的值在遇到存在列时得到赋值存在cell的值
					old = UUID.randomUUID().toString();
					
					col++;
					continue;
				}
				Cell cell = rowTemp.getCell(col);
				//处理不存在的cell
				if(cell == null){
					if(col - start > 1){
						CellRangeAddress cra =new CellRangeAddress(i,i,start,col-1);
						sheet.addMergedRegion(cra);
					}
					start = col;
					//保证老的值在遇到存在行时得到赋值存在cell的值
					old = UUID.randomUUID().toString();
					
					col++;
					continue;
				}
				
				String value = cell.getStringCellValue();
				if(!old.equals(value)){
					if(col - start > 1){
						CellRangeAddress cra =new CellRangeAddress(i,i,start,col-1);
						sheet.addMergedRegion(cra);
					}
					old = value;
					start = col;
				}
				if(col == endCol){
					if(col - start >= 1){
						CellRangeAddress cra =new CellRangeAddress(i,i,start,col);
						sheet.addMergedRegion(cra);
					}
				}
				col++;
			}
		}
		return wb;
	}
	public Workbook mergeCols(String name){
		return mergeCols(name,null,null,null,null);
	}
	
	/**
	 * 填充每行的数据
	 *
	 * @param dataRow   行对象
	 * @param cellIndex 单元格索引
	 * @param value     将被填入的值对象
	 */
	private void fillDataCell(Row dataRow, int cellIndex, Object value) {
		//创建一个单元格
		Cell cell = dataRow.createCell(cellIndex);
		cell.setCellStyle(this.cellStyle);
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
	 * 设置表格开始位置
	 * @param startRow 开始行 默认0 起始:0-
	 * @param startCol 开始列 默认0 起始:0-
	 */
	public void setStartLocation(Integer startRow, Short startCol){
		this.index = this.startRow = startRow;
		this.startCol = startCol;
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
	
	
	
	public Integer getStartRow() {
		return startRow;
	}
	
	public void setStartRow(Integer startRow) {
		this.startRow = startRow;
	}
	
	public Short getStartCol() {
		return startCol;
	}
	
	public void setStartCol(Short startCol) {
		this.startCol = startCol;
	}
	
	public CellStyle getCellStyle() {
		return cellStyle;
	}
	
	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	
	public Workbook getWorkbook(){
		return this.wb;
	}
	
	public Integer getLastRowNum(){
		return this.wb.getSheet(this.currentSheetName).getLastRowNum();
	}
	
	public Integer getLastRow() {
		return lastRow;
	}
	
	public Short getLastCol() {
		return lastCol;
	}
	
	public Integer getIndex() {
		return index;
	}
	
	public void setIndex(Integer index) {
		this.index = index;
	}
	
	public String getHeadSeparator() {
		return headSeparator;
	}
	
	public void setHeadSeparator(String headSeparator) {
		this.headSeparator = headSeparator;
	}
	
	public String getCurrentSheetName() {
		return currentSheetName;
	}
}
