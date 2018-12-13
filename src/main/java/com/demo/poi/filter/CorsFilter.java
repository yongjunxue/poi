package com.demo.poi.filter;

import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletResponse;

/**
 * 跨域访问
 * @author xueyongjun
 *
 */
public class CorsFilter implements Filter {

	@Override
	public void init(FilterConfig filterConfig) throws ServletException {
		
	}

	@Override
	public void doFilter(ServletRequest request, ServletResponse response,
			FilterChain chain) throws IOException, ServletException {
		((HttpServletResponse) response).setHeader("Access-Control-Allow-Origin", "*");
		((HttpServletResponse) response).addHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
		chain.doFilter(request, response);
	}

	@Override
	public void destroy() {
		
	}

}
