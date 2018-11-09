package com.demo.poi.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.demo.poi.exception.ComException;
import com.demo.poi.util.excel.part.Block;
import com.demo.poi.util.excel.part.Color;
import com.demo.poi.util.excel.part.ExcelModel;
import com.demo.poi.util.excel.part.ISheet;
import com.demo.poi.util.excel.part.SpecialSheet;
import com.demo.poi.util.excel.part.Style;

/**
 * 由于时间问题，工具类暂时借鉴网上已有的。后续会进行兼容性改进
 * @author xueyongjun
 *
 */
public class ExcelUtil {
	private final static String excel2003L = ".xls"; // 2003- 版本的excel
	private final static String excel2007U = ".xlsx"; // 2007+ 版本的excel
	
	@Deprecated  // TODO
	public static List<List<Object>> read(InputStream in, String fileName) throws Exception {
		List<List<Object>> list = null;
 
		// 创建Excel工作薄
		Workbook work = getWorkbook(in, fileName);
		if (null == work) {
			throw new Exception("创建Excel工作薄为空！");
		}
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
 
		list = new ArrayList<List<Object>>();
		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null) {
				continue;
			}
			// 遍历当前sheet中的所有行
			for (int j = sheet.getFirstRowNum(); j < sheet.getLastRowNum() + 1; j++) { // 这里的加一是因为下面的循环跳过取第一行表头的数据内容了
				row = sheet.getRow(j);
				if (row == null || row.getFirstCellNum() == j) {
					continue;
				}
				// 遍历所有的列
				List<Object> li = new ArrayList<Object>();
				for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
					cell = row.getCell(y);
					li.add(getCellValue(cell));
				}
				list.add(li);
			}
		}
		work.close();
		return list;
	}
	
	private static Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
		Workbook wb = null;
		String fileType = fileName.substring(fileName.lastIndexOf("."));
		if (excel2003L.equals(fileType)) {
			wb = new HSSFWorkbook(inStr); // 2003-
		} else if (excel2007U.equals(fileType)) {
			wb = new XSSFWorkbook(inStr); // 2007+
		} else {
			throw new Exception("解析的文件格式有误！");
		}
		return wb;
	}

	private static Object getCellValue(Cell cell) {
		Object value = null;
		DecimalFormat df = new DecimalFormat("0"); // 格式化number String字符
		SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd"); // 日期格式化
		DecimalFormat df2 = new DecimalFormat("0.00"); // 格式化数字
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				value = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if ("General".equals(cell.getCellStyle().getDataFormatString())) {
					value = df.format(cell.getNumericCellValue());
				} else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
					value = sdf.format(cell.getDateCellValue());
				} else {
					value = df2.format(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				value = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_BLANK:
				value = "";
				break;
			default:
				break;
		}
		return value;
	}

	
//-----导出-------------------------------------------------------------------
//	public static void export(List<List<Object>> list,
//			HttpServletResponse response) {
//		ExcelModel em=new ExcelModel();
//		com.demo.poi.util.excel.ExcelModel.Sheet sheet=em.createSheet();
//		Block b=sheet.createBlock();
//		b.setDataList(list);
//		export(em,response);
//	}
	
	/**
	 * 不建议使用这个，一个sheet中的行数大于65535时会报OOM错误，并且xls格式的用微软的办公软件打不开
	 * @param excel
	 * @param response
	 */
	public static void export(ExcelModel excel,HttpServletResponse response){
		response.setContentType("application/binary;charset=UTF-8");// new String(filename.getBytes("gbk")
		String name="";
		try {
			name=new String(excel.getExcelName().getBytes("gbk"),"iso8859-1");
		} catch (UnsupportedEncodingException e1) {
			e1.printStackTrace();
			throw new ComException("字符转换异常");
		}
		response.setHeader("Content-disposition", "attachment; filename=" + name + ".xls");
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		for(ISheet iSheet : excel.getSheets()){
			if(iSheet instanceof com.demo.poi.util.excel.part.Sheet){
				com.demo.poi.util.excel.part.Sheet mySheet=(com.demo.poi.util.excel.part.Sheet)iSheet;
				HSSFSheet sheet= workbook.createSheet(mySheet.getSheetName());
				
				List<Block> blocks=mySheet.getBlocks();
				for(Block block : blocks){
					setTitle(workbook,sheet,block,mySheet);
					
					setHeader(workbook,sheet,block,mySheet);
					
					setData(workbook,sheet,block,mySheet);
					
					setFooter(workbook,sheet,block,mySheet);
					
					setBlockSpace(workbook,sheet,mySheet);
					
				}
				setCellWidth(sheet,mySheet);
			}else if(iSheet instanceof SpecialSheet){
				SpecialSheet mySheet=(SpecialSheet)iSheet;
				setData(workbook,mySheet);
			}
		}
		
		try {
			ServletOutputStream out=response.getOutputStream();
			workbook.write(out);
			out.flush();
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
			throw new ComException("导出失败");
		}
	}

	public static void exportXlsx(ExcelModel excel,HttpServletResponse response){
		response.setContentType("application/binary;charset=UTF-8");// new String(filename.getBytes("gbk")
		String name="";
		try {
			name=new String(excel.getExcelName().getBytes("gbk"),"iso8859-1");
		} catch (UnsupportedEncodingException e1) {
			e1.printStackTrace();
			throw new ComException("字符转换异常");
		}
		response.setHeader("Content-disposition", "attachment; filename=" + name + ".xlsx");
		
		SXSSFWorkbook workbook = new SXSSFWorkbook();
		for(com.demo.poi.util.excel.part.ISheet iSheet : excel.getSheets()){
			if(iSheet instanceof com.demo.poi.util.excel.part.Sheet){
				com.demo.poi.util.excel.part.Sheet mySheet=(com.demo.poi.util.excel.part.Sheet)iSheet;
				SXSSFSheet sheet= workbook.createSheet(mySheet.getSheetName());
				
				List<Block> blocks=mySheet.getBlocks();
				for(Block block : blocks){
					setTitleXlsx(workbook,sheet,block,mySheet);
					
					setHeaderXlsx(workbook,sheet,block,mySheet);
					
					setDataXlsx(workbook,sheet,block,mySheet);
					
					setFooterXlsx(workbook,sheet,block,mySheet);
					
					setBlockSpaceXlsx(workbook,sheet,mySheet);
					
				}
				setCellWidthXlsx(sheet,mySheet);
			}else if(iSheet instanceof SpecialSheet){
				SpecialSheet mySheet=(SpecialSheet)iSheet;
				setDataXlsx(workbook,mySheet);
			}
		}
		
		try {
			ServletOutputStream out=response.getOutputStream();
			workbook.write(out);
			out.flush();
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
			throw new ComException("导出失败");
		}
	}
	
	private static void setData(HSSFWorkbook workbook, SpecialSheet mySheet) {
		HSSFSheet sheet= workbook.createSheet(mySheet.getSheetName());
		List<com.demo.poi.util.excel.part.Row> rows=mySheet.getRows();
		if(rows != null){
			int rowNo=0;
			for(com.demo.poi.util.excel.part.Row myRow : rows){
				HSSFRow row=sheet.createRow(rowNo);
				int cellNo=0;
				for(com.demo.poi.util.excel.part.Cell myCell : myRow.getCells()){
					
					HSSFCellStyle cellStyle=null;
					if(myCell.getStyle() != null){
						cellStyle= createStyle(workbook, myCell.getStyle());
					}

					Cell cell=row.createCell(cellNo);
					cell.setCellValue(myCell.getValue());
					cell.setCellStyle(cellStyle);
					
					sheet.setColumnWidth(cellNo, myCell.getWidth());
					if(myCell.isMerge()){
						CellRangeAddress cra = new CellRangeAddress(rowNo, rowNo, myCell.getFirstCol(), myCell.getLastCol()); 
						sheet.addMergedRegion(cra);
						
						
						for(int i=1;i<=(myCell.getLastCol()-myCell.getFirstCol());i++){
							cellNo++;
						}
					}
					
					cellNo++;
				}
				rowNo++;
			}
		}
	}
	
	private static void setDataXlsx(SXSSFWorkbook workbook, SpecialSheet mySheet) {
		SXSSFSheet sheet= workbook.createSheet(mySheet.getSheetName());
		List<com.demo.poi.util.excel.part.Row> rows=mySheet.getRows();
		if(rows != null){
			int rowNo=0;
			for(com.demo.poi.util.excel.part.Row myRow : rows){
				SXSSFRow row=sheet.createRow(rowNo);
				int cellNo=0;
				for(com.demo.poi.util.excel.part.Cell myCell : myRow.getCells()){
					
					CellStyle cellStyle=null;
					if(myCell.getStyle() != null){
						cellStyle= createStyleXlsx(workbook, myCell.getStyle());
					}

					Cell cell=row.createCell(cellNo);
					cell.setCellValue(myCell.getValue());
					cell.setCellStyle(cellStyle);
					
					sheet.setColumnWidth(cellNo, myCell.getWidth());
					if(myCell.isMerge()){
						CellRangeAddress cra = new CellRangeAddress(rowNo, rowNo, myCell.getFirstCol(), myCell.getLastCol()); 
						sheet.addMergedRegion(cra);
						
						
						for(int i=1;i<=(myCell.getLastCol()-myCell.getFirstCol());i++){
							cellNo++;
						}
					}
					
					cellNo++;
				}
				rowNo++;
			}
		}
	}
	
	private static void setCellWidth(HSSFSheet sheet,
			com.demo.poi.util.excel.part.Sheet mySheet) {
		for(int i=0;i<mySheet.getMaxColCount();i++){
			sheet.autoSizeColumn(i);
		}
	}
	
	private static void setCellWidthXlsx(SXSSFSheet sheet,
			com.demo.poi.util.excel.part.Sheet mySheet) {
		sheet.trackAllColumnsForAutoSizing();
		for(int i=0;i<mySheet.getMaxColCount();i++){
			sheet.autoSizeColumn(i);
		}
	}

	private static void setBlockSpace(HSSFWorkbook workbook, HSSFSheet sheet,
			com.demo.poi.util.excel.part.Sheet mySheet) {
		for(int i=1;i<=mySheet.getBlockSpace();i++){
			sheet.createRow(mySheet.nextRowNo());
		}
	}
	
	private static void setBlockSpaceXlsx(SXSSFWorkbook workbook, SXSSFSheet sheet,
			com.demo.poi.util.excel.part.Sheet mySheet) {
		for(int i=1;i<=mySheet.getBlockSpace();i++){
			sheet.createRow(mySheet.nextRowNo());
		}
	}

	private static void setFooter(HSSFWorkbook workbook, HSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		com.demo.poi.util.excel.part.Cell myFooter=block.getFooter();
		if(myFooter != null){
			HSSFCellStyle defaultStyle=getDefaultStyle(workbook);
			
			HSSFCellStyle footerStyle=null;
			if(block.getFooterStyle() != null){
				footerStyle= createStyle(workbook, block.getFooterStyle());
				footerStyle.setAlignment(HorizontalAlignment.LEFT);
			}
			
			int cellIndex=0;
			HSSFRow row=sheet.createRow(mySheet.nextRowNo());
			
			HSSFCell footer=row.createCell(cellIndex);
			footer.setCellValue(myFooter.getValue());
			
			if(block.getColCount()>1){
				//合并单元格
				CellRangeAddress cra = new CellRangeAddress(mySheet.getCurrentRowNo(), mySheet.getCurrentRowNo(), 0, block.getColCount()-1); 
				sheet.addMergedRegion(cra);
			}
			
			HSSFCellStyle cellStyle=null;
			if(myFooter.getStyle() != null){
				cellStyle=createStyle(workbook, myFooter.getStyle());
				footerStyle.setAlignment(HorizontalAlignment.LEFT);
			}
			
			if(footerStyle != null){
				footer.setCellStyle(footerStyle);
			}else{
				footer.setCellStyle(defaultStyle);
			}
			if(cellStyle != null){
				footer.setCellStyle(cellStyle);
			}
		}
	}
	
	private static void setFooterXlsx(SXSSFWorkbook workbook, SXSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		com.demo.poi.util.excel.part.Cell myFooter=block.getFooter();
		if(myFooter != null){
			CellStyle defaultStyle=getDefaultStyleXlsx(workbook);
			
			CellStyle footerStyle=null;
			if(block.getFooterStyle() != null){
				footerStyle= createStyleXlsx(workbook, block.getFooterStyle());
				footerStyle.setAlignment(HorizontalAlignment.LEFT);
			}
			
			int cellIndex=0;
			SXSSFRow row=sheet.createRow(mySheet.nextRowNo());
			
			SXSSFCell footer=row.createCell(cellIndex);
			footer.setCellValue(myFooter.getValue());
			
			if(block.getColCount()>1){
				//合并单元格
				CellRangeAddress cra = new CellRangeAddress(mySheet.getCurrentRowNo(), mySheet.getCurrentRowNo(), 0, block.getColCount()-1); 
				sheet.addMergedRegion(cra);
			}
			
			CellStyle cellStyle=null;
			if(myFooter.getStyle() != null){
				cellStyle=createStyleXlsx(workbook, myFooter.getStyle());
				footerStyle.setAlignment(HorizontalAlignment.LEFT);
			}
			
			if(footerStyle != null){
				footer.setCellStyle(footerStyle);
			}else{
				footer.setCellStyle(defaultStyle);
			}
			if(cellStyle != null){
				footer.setCellStyle(cellStyle);
			}
		}
	}
	
	private static void setData(HSSFWorkbook workbook, HSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		
		if(block.getRows()!=null){
			int rowNo=1; //序号
			
			HSSFCellStyle defaultStyle=getDefaultStyle(workbook);
			
			HSSFCellStyle allRowStyle=null;//所有行样式
			if(block.getRowsStyle() != null){
				allRowStyle=createStyle(workbook, block.getRowsStyle());
			}
			
			for(com.demo.poi.util.excel.part.Row myRow : block.getRows()){
				HSSFCellStyle rowStyle=null;//单行样式
				if(myRow.getStyle() != null){
					rowStyle=createStyle(workbook, myRow.getStyle());
				}

				HSSFRow row=sheet.createRow(mySheet.nextRowNo());
				int cellIndex=0;
				
				if(block.isShowRowNo()){
					HSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowNo++);
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
						if(rowStyle != null){
							cell.setCellStyle(rowStyle);
						}
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
				
				for(com.demo.poi.util.excel.part.Cell myCell : myRow.getCells()){
					HSSFCellStyle cellStyle=null;//单元格样式
					if(myCell.getStyle() != null){
						cellStyle=createStyle(workbook, myCell.getStyle());
					}
					
					HSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(myCell.getValue());
					
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
					if(rowStyle != null){
						cell.setCellStyle(rowStyle);
					}
					if(cellStyle != null){
						cell.setCellStyle(cellStyle);
					}
					
					
				}
			}
		}else if(block.getData() !=null){
			HSSFCellStyle defaultStyle= getDefaultStyle(workbook);
			
			HSSFCellStyle allRowStyle=null;//所有行样式
			if(block.getRowsStyle() != null){
				allRowStyle=createStyle(workbook, block.getRowsStyle());
			}
			
			int rowNo=1;
			for(Map rowMap : block.getData()){
				HSSFRow row=sheet.createRow(mySheet.nextRowNo());
				int cellIndex=0;
				
				if(block.isShowRowNo()){
					HSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowNo++);
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
				
				for(String key : block.getHeaderKeys()){
					HSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowMap.get(key).toString());
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
			}
		}
	}
	
	private static void setDataXlsx(SXSSFWorkbook workbook, SXSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		
		if(block.getRows()!=null){
			int rowNo=1; //序号
			
			CellStyle defaultStyle=getDefaultStyleXlsx(workbook);
			
			CellStyle allRowStyle=null;//所有行样式
			if(block.getRowsStyle() != null){
				allRowStyle=createStyleXlsx(workbook, block.getRowsStyle());
			}
			
			for(com.demo.poi.util.excel.part.Row myRow : block.getRows()){
				CellStyle rowStyle=null;//单行样式
				if(myRow.getStyle() != null){
					rowStyle=createStyleXlsx(workbook, myRow.getStyle());
				}

				SXSSFRow row=sheet.createRow(mySheet.nextRowNo());
				int cellIndex=0;
				
				if(block.isShowRowNo()){
					SXSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowNo++);
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
						if(rowStyle != null){
							cell.setCellStyle(rowStyle);
						}
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
				
				for(com.demo.poi.util.excel.part.Cell myCell : myRow.getCells()){
					CellStyle cellStyle=null;//单元格样式
					if(myCell.getStyle() != null){
						cellStyle=createStyleXlsx(workbook, myCell.getStyle());
					}
					
					SXSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(myCell.getValue());
					
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
					if(rowStyle != null){
						cell.setCellStyle(rowStyle);
					}
					if(cellStyle != null){
						cell.setCellStyle(cellStyle);
					}
					
					
				}
			}
		}else if(block.getData() !=null){
			CellStyle defaultStyle= getDefaultStyleXlsx(workbook);
			
			CellStyle allRowStyle=null;//所有行样式
			if(block.getRowsStyle() != null){
				allRowStyle=createStyleXlsx(workbook, block.getRowsStyle());
			}
			
			int rowNo=1;
			for(Map rowMap : block.getData()){
				SXSSFRow row=sheet.createRow(mySheet.nextRowNo());
				int cellIndex=0;
				
				if(block.isShowRowNo()){
					SXSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowNo++);
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
				
				for(String key : block.getHeaderKeys()){
					SXSSFCell cell=row.createCell(cellIndex++);
					cell.setCellValue(rowMap.get(key).toString());
					if(allRowStyle != null){
						cell.setCellStyle(allRowStyle);
					}else{
						cell.setCellStyle(defaultStyle);
					}
				}
			}
		}
	}
	
	private static HSSFCellStyle getDefaultStyle(HSSFWorkbook workbook) {
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);//默认居中
		return cellStyle;
	}
	
	private static CellStyle getDefaultStyleXlsx(SXSSFWorkbook workbook) {
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);//默认居中
		return cellStyle;
	}
	
	private static void setHeader(HSSFWorkbook workbook, HSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		if(block.getHeader()!= null){
			HSSFCellStyle defaultStyle=getDefaultStyle(workbook);
			
			int cellIndex=0;
			HSSFRow row=sheet.createRow(mySheet.nextRowNo());
			
			HSSFCellStyle headerStyle=null;
			if(block.getHeaderStyle() != null){
				headerStyle= createStyle(workbook, block.getHeaderStyle());
			}
			
			if(block.isShowRowNo()){
				HSSFCell cell=row.createCell(cellIndex++);
				cell.setCellValue(block.getRowNoName());
				if(headerStyle !=null){
					cell.setCellStyle(headerStyle);
				}else{
					cell.setCellStyle(defaultStyle);
				}
			}
			for(com.demo.poi.util.excel.part.Cell myHeader : block.getHeader()){
				HSSFCell header=row.createCell(cellIndex++);
				header.setCellValue(myHeader.getValue());
				
				HSSFCellStyle cellStyle=null;
				if(myHeader.getStyle() !=null){
					cellStyle=createStyle(workbook, myHeader.getStyle());
				}
				
				if(headerStyle != null){
					header.setCellStyle(headerStyle);
				}else{
					header.setCellStyle(defaultStyle);
				}
				if(cellStyle != null){
					header.setCellStyle(cellStyle);
				}
			}
		}
	}
	
	private static void setHeaderXlsx(SXSSFWorkbook workbook, SXSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		if(block.getHeader()!= null){
			CellStyle defaultStyle=getDefaultStyleXlsx(workbook);
			
			int cellIndex=0;
			SXSSFRow row=sheet.createRow(mySheet.nextRowNo());
			
			CellStyle headerStyle=null;
			if(block.getHeaderStyle() != null){
				headerStyle= createStyleXlsx(workbook, block.getHeaderStyle());
			}
			
			if(block.isShowRowNo()){
				SXSSFCell cell=row.createCell(cellIndex++);
				cell.setCellValue(block.getRowNoName());
				if(headerStyle !=null){
					cell.setCellStyle(headerStyle);
				}else{
					cell.setCellStyle(defaultStyle);
				}
			}
			for(com.demo.poi.util.excel.part.Cell myHeader : block.getHeader()){
				SXSSFCell header=row.createCell(cellIndex++);
				header.setCellValue(myHeader.getValue());
				
				CellStyle cellStyle=null;
				if(myHeader.getStyle() !=null){
					cellStyle=createStyleXlsx(workbook, myHeader.getStyle());
				}
				
				if(headerStyle != null){
					header.setCellStyle(headerStyle);
				}else{
					header.setCellStyle(defaultStyle);
				}
				if(cellStyle != null){
					header.setCellStyle(cellStyle);
				}
			}
		}
	}
	
	private static void setTitle(HSSFWorkbook workbook, HSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		com.demo.poi.util.excel.part.Cell myTitle=block.getTitle();
		HSSFRow row=null;
		int cellIndex=0;
		if(myTitle != null){
			HSSFCellStyle defaultStyle=getDefaultStyle(workbook);
			
			row=sheet.createRow(mySheet.nextRowNo());
			HSSFCell title=row.createCell(cellIndex);
			
			//样式TODO
			HSSFCellStyle titleStyle=null;
			if(block.getTitleStyle() != null){
				titleStyle=createStyle(workbook, block.getTitleStyle());
			}
			
			if(block.getColCount()>1){
				//合并单元格
				CellRangeAddress cra = new CellRangeAddress(mySheet.getCurrentRowNo(), mySheet.getCurrentRowNo(), 0, block.getColCount()-1); 
				sheet.addMergedRegion(cra);
			}
			
			title.setCellValue(myTitle.getValue());
			if(titleStyle != null){
				title.setCellStyle(titleStyle);
			}else{
				title.setCellStyle(defaultStyle);
			}
		}
	}
	
	private static void setTitleXlsx(SXSSFWorkbook workbook, SXSSFSheet sheet,
			Block block, com.demo.poi.util.excel.part.Sheet mySheet) {
		com.demo.poi.util.excel.part.Cell myTitle=block.getTitle();
		SXSSFRow row=null;
		int cellIndex=0;
		if(myTitle != null){
			CellStyle defaultStyle=getDefaultStyleXlsx(workbook);
			
			row=sheet.createRow(mySheet.nextRowNo());
			SXSSFCell title=row.createCell(cellIndex);
			
			//样式TODO
			CellStyle titleStyle=null;
			if(block.getTitleStyle() != null){
				titleStyle=createStyleXlsx(workbook, block.getTitleStyle());
			}
			
			if(block.getColCount()>1){
				//合并单元格
				CellRangeAddress cra = new CellRangeAddress(mySheet.getCurrentRowNo(), mySheet.getCurrentRowNo(), 0, block.getColCount()-1); 
				sheet.addMergedRegion(cra);
			}
			
			title.setCellValue(myTitle.getValue());
			if(titleStyle != null){
				title.setCellStyle(titleStyle);
			}else{
				title.setCellStyle(defaultStyle);
			}
		}
	}
	
	private static HSSFCellStyle createStyle(HSSFWorkbook workbook,
			Style style) {
		
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		
		HSSFFont font = workbook.createFont();
		
		if(style!=null){
			if(style.getBackColor() != null){
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cellStyle.setFillForegroundColor(style.getBackColor());
				setBorderColor(cellStyle);
			}
			if(style.getFontColor() != null){
				font.setColor(style.getFontColor());
			}
			if(style.getSize()!=null){
				font.setFontHeightInPoints(style.getSize());
			}
			if(style.getBold() != null){
				font.setBold(style.getBold());
			}
			if(style.getAlign() != null){
				cellStyle.setAlignment(style.getAlign());//默认居中
			}
			if(style.getFontName() != null){
				font.setFontName(style.getFontName());
			}
		}
		cellStyle.setFont(font);
		
		return cellStyle;
	}
	
	private static CellStyle createStyleXlsx(SXSSFWorkbook workbook,
			Style style) {
		
		CellStyle cellStyle = workbook.createCellStyle();
		
		Font font = workbook.createFont();
		
		if(style!=null){
			if(style.getBackColor() != null){
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				cellStyle.setFillForegroundColor(style.getBackColor());
				setBorderColor(cellStyle);
			}
			if(style.getFontColor() != null){
				font.setColor(style.getFontColor());
			}
			if(style.getSize()!=null){
				font.setFontHeightInPoints(style.getSize());
			}
			if(style.getBold() != null){
				font.setBold(style.getBold());
			}
			if(style.getAlign() != null){
				cellStyle.setAlignment(style.getAlign());//默认居中
			}
			if(style.getFontName() != null){
				font.setFontName(style.getFontName());
			}
		}
		cellStyle.setFont(font);
		
		return cellStyle;
	}


	private static void setBorderColor(Object cellStyle_obj) {
		if(cellStyle_obj instanceof HSSFCellStyle){
			HSSFCellStyle cellStyle=(HSSFCellStyle)cellStyle_obj;
			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderRight(BorderStyle.THIN);
			
			cellStyle.setTopBorderColor(Color.BLACK);
			cellStyle.setBottomBorderColor(Color.BLACK);
			cellStyle.setLeftBorderColor(Color.BLACK);
			cellStyle.setRightBorderColor(Color.BLACK);
		}else if(cellStyle_obj instanceof XSSFCellStyle){
			XSSFCellStyle cellStyle=(XSSFCellStyle)cellStyle_obj;
			cellStyle.setBorderTop(BorderStyle.THIN);
			cellStyle.setBorderBottom(BorderStyle.THIN);
			cellStyle.setBorderLeft(BorderStyle.THIN);
			cellStyle.setBorderRight(BorderStyle.THIN);
			
			cellStyle.setTopBorderColor(Color.BLACK);
			cellStyle.setBottomBorderColor(Color.BLACK);
			cellStyle.setLeftBorderColor(Color.BLACK);
			cellStyle.setRightBorderColor(Color.BLACK);
		}
	}

}
