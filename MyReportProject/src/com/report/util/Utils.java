package com.report.util;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

public class Utils {

	private static final String SPRING_MODULE_PATH = "Spring-Module.xml";
	private static final String CONFIG_PATH = "config.properties";
	public static final ApplicationContext APP_CONTEXT = new ClassPathXmlApplicationContext(SPRING_MODULE_PATH);
	
	/*public static ApplicationContext getApplicationContext(){
		return new ClassPathXmlApplicationContext(SPRING_MODULE_PATH);
	}*/
	
	public static String getProperties(String key){
		InputStream inputStream = null;
		String result = "";
		try {
			Properties prop = new Properties();
			inputStream = Utils.class.getClassLoader().getResourceAsStream(CONFIG_PATH);
 
			if (inputStream != null) {
				prop.load(inputStream);
			} else {
				throw new FileNotFoundException("property file '" + CONFIG_PATH + "' not found in the classpath");
			}
 
			return prop.getProperty(key);
		} catch (Exception e) {
			System.out.println("Exception: " + e);
		} finally {
			try {
				inputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return result;
	}
	
}
