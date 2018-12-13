package com.demo.poi.angular;

import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.demo.poi.service.ExportService;

@Controller
public class Angular6DemoController {
	
	JSONArray exams=new JSONArray();

	
	@Autowired
	ExportService exportService;
	
	@RequestMapping(value = "/exams")
	@ResponseBody
    public String getHeroes(@RequestBody String param,HttpServletResponse response) throws Exception{
		response.setHeader("Access-Control-Allow-Origin", "*");
		System.out.println(param);
		JSONObject jp=JSONObject.fromObject(param);
		JSONObject paper=jp.getJSONObject("paper");
		int page=paper.getInt("page");
		int pageSize=paper.getInt("pageSize");
		int total=exams.size();
		
		int startIndex=(page-1)*pageSize;
		int endIndex=startIndex+pageSize-1;
		JSONArray array=new JSONArray();
		for(int i=startIndex;i<=endIndex;i++){
			if(i>=exams.size()){
				break;
			}
			array.add(exams.get(i));
		}
		
		JSONObject result=new JSONObject();
		paper.put("page", page);
		paper.put("pageSize", pageSize);
		paper.put("total", total);
		result.put("paper", paper);
		result.put("examList", array);
		
		return result.toString();
	}
	
	@RequestMapping(value = "/exam/pass")
	@ResponseBody
    public String pass(@RequestBody String param,HttpServletResponse response) throws Exception{
		response.setHeader("Access-Control-Allow-Origin", "*");
		System.out.println(param);
		JSONObject item=new JSONObject();
		item.put("pass", param);
		return item.toString();
	}
	
	@RequestMapping(value = "/exam/reject")
	@ResponseBody
    public String reject(@RequestBody String param,HttpServletResponse response) throws Exception{
		response.setHeader("Access-Control-Allow-Origin", "*");
		System.out.println(param);
		JSONObject item=new JSONObject();
		item.put("reject", param);
		return item.toString();
	}
	
	{
		JSONObject item=new JSONObject();
		item.put("id", 1);
		item.put("examName", "月考");
		item.put("examType", "1000");
		item.put("score", "120");
		item.put("verifyStatus", "1");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 2);
		item.put("examName", "周考");
		item.put("examType", "2000");
		item.put("score", "150");
		item.put("verifyStatus", "2");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 3);
		item.put("examName", "季度考");
		item.put("examType", "1000");
		item.put("score", "100");
		item.put("verifyStatus", "2");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 4);
		item.put("examName", "期末考");
		item.put("examType", "1000");
		item.put("score", "150");
		item.put("verifyStatus", "1");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 5);
		item.put("examName", "周考");
		item.put("examType", "2000");
		item.put("score", "100");
		item.put("verifyStatus", "1");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 6);
		item.put("examName", "周考");
		item.put("examType", "1000");
		item.put("score", "130");
		item.put("verifyStatus", "2");
		exams.add(item);
		
		item=new JSONObject();
		item.put("id", 7);
		item.put("examName", "升学考试");
		item.put("examType", "2000");
		item.put("score", "130");
		item.put("verifyStatus", "2");
		exams.add(item);
	}
}
