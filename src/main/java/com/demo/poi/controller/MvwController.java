package com.demo.poi.controller;

import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONObject;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.demo.poi.service.ExportService;

@Controller
public class MvwController {
	
	String name="xyj";
	
	String pwd="111111";
	
	String token;
	
	@Autowired
	ExportService exportService;
	
	@RequestMapping(value = "/login")
	@ResponseBody
    public String export(@RequestBody String body,HttpServletResponse response) throws Exception{
		JSONObject result=new JSONObject();
		
		response.setHeader("Access-Control-Allow-Origin", "*");
		response.addHeader("Access-Control-Allow-Headers", "Content-Type");
		response.addHeader("Access-Control-Allow-Credentials", "true");
        response.addHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE,OPTIONS");
		JSONObject param=JSONObject.fromObject(URLDecoder.decode(body));
		if(name.equals(param.get("name")) && pwd.equals(param.get("pwd"))){
			result.put("result", "true");
			token=UUID.randomUUID().toString().replaceAll("-", "");
			result.put("token", token);
			result.put("msg", "操作成功");
		}else{
			result.put("result", "false");
			result.put("msg", "用户名或密码错误");
		}
		
		System.out.println(URLEncoder.encode(result.toString()));
		return URLEncoder.encode(result.toString());
	}
	
	@RequestMapping(value = "/delExam")
	@ResponseBody
    public String delExam(@RequestBody String body,HttpServletRequest req,HttpServletResponse response) throws Exception{
		JSONObject re=new JSONObject();
		response.setHeader("Access-Control-Allow-Origin", "*");
        String token=req.getParameter("MYTOKEN");
        System.out.println("前端传来的token:"+token);
        System.out.println(JSONObject.fromObject(URLDecoder.decode(body)));
        if(this.token.equals(token)){
        	re.put("result", "true");
        	re.put("msg", "操作成功");
        	re.put("token", this.token);
        }else{
        	re.put("result", "false");
        	re.put("msg", "token失效，请重新登陆");
        }
		
		return URLEncoder.encode(re.toString());
	}
	
	@RequestMapping(value = "/verify/{id}")
	@ResponseBody
    public String verify(@PathVariable("id") String id,HttpServletRequest req,HttpServletResponse response) throws Exception{
		JSONObject re=new JSONObject();
		response.setHeader("Access-Control-Allow-Origin", "*");
        String token=req.getParameter("MYTOKEN");
        System.out.println("前端传来的token:"+token);
//        System.out.println(JSONObject.fromObject(URLDecoder.decode(body)));
        System.out.println("examId:"+id);
        if(this.token.equals(token)){
        	re.put("result", "true");
        	re.put("msg", "操作成功");
        	re.put("token", this.token);
        }else{
        	re.put("result", "false");
        	re.put("msg", "token失效，请重新登陆");
        }
		
		return URLEncoder.encode(re.toString());
	}
}
