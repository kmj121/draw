package com.example.draw.controller;

import com.example.draw.dto.ReportData;
import com.example.draw.service.WordService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

@RestController
public class WordController {

    @Autowired
    public WordService wordService;

    @PostMapping(value = "/aaa")
    public void aaa(@RequestBody ReportData data) throws Exception {
        System.out.println("aaa");
        wordService.test(data);
    }
}
