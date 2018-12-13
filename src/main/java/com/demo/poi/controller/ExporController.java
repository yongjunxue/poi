package com.demo.poi.controller;

import java.io.PrintWriter;
import java.net.URLDecoder;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONObject;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.demo.poi.exception.ComException;
import com.demo.poi.service.ExportService;
import com.demo.poi.util.ExcelUtil;
import com.demo.poi.util.excel.part.ExcelModel;

@Controller
public class ExporController {
	
	@Autowired
	ExportService exportService;
	
	@RequestMapping(value = "/export")
    public String export(HttpServletRequest request,HttpServletResponse response) throws Exception{
		response.setHeader("Access-Control-Allow-Origin", "*");
		
		ExcelModel em=null;
		try{
			em=exportService.getExcelModel();
			ExcelUtil.export(em, response);
		}catch(ComException e){
			e.printStackTrace();
			request.setCharacterEncoding("GBK");
            response.setCharacterEncoding("GBK");
            PrintWriter print = response.getWriter();
            print.write(e.getMessage());
		}catch(Exception e){
			e.printStackTrace();
			request.setCharacterEncoding("GBK");
            response.setCharacterEncoding("GBK");
            PrintWriter print = response.getWriter();
            print.write("导出异常");
		}
		return null;
	}
}
