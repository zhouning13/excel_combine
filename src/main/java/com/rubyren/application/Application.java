package com.rubyren.application;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletComponentScan;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableAutoConfiguration
@ComponentScan
@SpringBootApplication
@ComponentScan(basePackages = "com.rubyren")
@ServletComponentScan(basePackages = "com.rubyren")
public class Application {

	public static void main(String[] args) {
		SpringApplication.run(Application.class, args);
	}

}
