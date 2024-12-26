package com.example.demo;



import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Override
    public void addCorsMappings(CorsRegistry registry) {
        registry.addMapping("/**") // Applies to all endpoints
                .allowedOrigins("http://localhost:63342") // Your frontend origin
                .allowedMethods("GET", "POST", "PUT", "DELETE") // HTTP methods allowed
                .allowCredentials(true) // Allow cookies or authentication headers
                .maxAge(3600); // Cache the CORS configuration for 1 hour
    }
}
