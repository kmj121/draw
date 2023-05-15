package com.example.draw;

import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@Slf4j
public class DrawApplication {

    public static void main(String[] args) {
        SpringApplication.run(DrawApplication.class, args);
    }

}
