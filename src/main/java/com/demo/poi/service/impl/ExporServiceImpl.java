package com.demo.poi.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import net.sf.json.JSONObject;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.stereotype.Service;

import com.demo.poi.service.ExportService;
import com.demo.poi.util.excel.part.Block;
import com.demo.poi.util.excel.part.Cell;
import com.demo.poi.util.excel.part.Color;
import com.demo.poi.util.excel.part.ExcelModel;
import com.demo.poi.util.excel.part.Row;
import com.demo.poi.util.excel.part.Sheet;
import com.demo.poi.util.excel.part.SpecialSheet;
import com.demo.poi.util.excel.part.Style;

@Service
public class ExporServiceImpl implements ExportService{

	@Override
	public ExcelModel getExcelModel() {
ExcelModel em=new ExcelModel("导出测试");
		
		//-------导出1，数据区统一使用一种样式----------
		List<Map> list = getTestData();
        Sheet s= em.createSheet();
        s.setBlockSpace(3);
        Block b=s.createBlock();
        b.setTitle("aaa").setTitleStyle((new Style().setFontColor(Color.GREEN)));
        b.setHeader("名字","手机号","身份证");
        b.setHeaderKeys("name","phonenumber","idcard");
        b.setData(list).setRowsStyle((new Style().setFontColor(Color.LIGHT_ORANGE)));//****************************
        b.setFooter("我是表尾");
        b.setShowRowNo(true);
        
        //-------导出2，样式可以具体到单元格-----------
        JSONObject arguments=new JSONObject();
        Block b2=s.createBlock();
        b2.setTitle("bbb").setTitleStyle((new Style().setFontColor(Color.GREEN)));
        b2.setHeader("名字","手机号","身份证","邮箱").setHeaderStyle((new Style().setFontColor(Color.RED)).setBold(true));
        b2.setHeaderKeys("name","phonenumber","idcard","email");
        b2.setRowsStyle((new Style().setBackColor(Color.LIGHT_ORANGE)));
        setRows(b2,list);//可以设置某一行或者某一个单元格的样式
        b2.setShowRowNo(true);//显示序号
        
      //-------导出3.无规则数据的导出-----------
        SpecialSheet ss=em.createSpecialSheet("合并单元格测试");
        Map m=new HashMap<>();
        m.put("name", "试卷名称：ksm和纠错_2018_9_26_16_49");
        m.put("date", "考试时间：2018-09-26 16:49:33----2018-09-27 16:49:35");
        m.put("userName", "出卷人：赵瑞丽超级管理员");
        m.put("timeLimit", "考试时长：22");
        m.put("totalScore", "总分：60");
        m.put("count", "题数：3");
        setSpecialSheet(ss,m);
		
		return em;
	}

	private void setRows(Block b, List<Map> mList) {
		List<Row> rs=new ArrayList<>();
		for(Map rowMap : mList){
			Row r=new Row();
			for(String key : b.getHeaderKeys()){
				if(rowMap.get(key) != null){
					Cell c=r.createCell(rowMap.get(key).toString());
					if(key.equals("name")){
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("2")){
							r.setStyle((new Style().setBackColor(Color.BLUE)));//如果名字中包含“2”,则将整行设置为这种样式
						}
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("8")){
							c.setStyle(new Style().setBackColor(Color.LIGHT_CORNFLOWER_BLUE).setSize((short) 15));//如果名字中包含“8”,则将这个单元格设为这种样式
						}
					}else if(key.equals("code")){
						if(rowMap.get(key)!=null && rowMap.get(key).toString().contains("34")){
							c.setStyle(new Style().setBackColor(Color.LIGHT_CORNFLOWER_BLUE).setBold(true));
						}
					}
				}
				
			}
			rs.add(r);
		}
		b.setRows(rs);
	}
	
	private void setSpecialSheet(SpecialSheet ss, Map m) {
		List<Row> rs=new ArrayList<>();
		Row r=new Row();
		r.createCell("标题").setMergeRange(0, 3).setWidth(50000).setStyle(new Style().setSize((short) 20).setBold(true).setAlign(HorizontalAlignment.CENTER));
		rs.add(r);
		r=new Row();
		short size=12;
		
		Style style=new Style().setSize(size ).setAlign(HorizontalAlignment.LEFT);
		r.createCell(m.get("name").toString()).setMergeRange(0, 1).setWidth(15000).setStyle(style);
		r.createCell(m.get("date").toString()).setMergeRange(2, 3).setWidth(15000).setStyle(style);
		rs.add(r);
		r=new Row();
		r.createCell(m.get("userName").toString()).setWidth(10000).setStyle(style );
		r.createCell(m.get("timeLimit").toString()).setWidth(10000).setStyle(style);
		r.createCell(m.get("totalScore").toString()).setWidth(10000).setStyle(style );
		r.createCell(m.get("count").toString()).setWidth(10000).setStyle(style);
		rs.add(r);
		ss.setRows(rs);
	}

	private List<Map> getTestData() {
		List<Map> list =new ArrayList<>();
		Map item=new HashMap<>();
		item.put("name", "张三");//"name","phonenumber","idcard","email"
		item.put("phonenumber", "11111111111");
		item.put("idcard", "1");
		item.put("email", "123@qq.com");
		list.add(item);
		
		item=new HashMap<>();
		item.put("name", "李四");//"name","phonenumber","idcard","email"
		item.put("phonenumber", "22222222222");
		item.put("idcard", "2");
		item.put("email", "22@qq.com");
		list.add(item);
		
		return list;
	}

}
