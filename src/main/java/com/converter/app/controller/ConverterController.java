package com.converter.app.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.converter.app.service.ConverterService;
import org.springframework.web.bind.annotation.GetMapping;


@RestController
@RequestMapping("/api")
public class ConverterController {

	@Autowired
	private ConverterService service;
	
	
	@GetMapping("/convert")
	public ResponseEntity<byte[]> getMethodName() throws Exception {
	    byte[] docxContent = service.convertToDocx();  

	    HttpHeaders headers = new HttpHeaders();
	    headers.add("Content-Disposition", "attachment; filename=mrfForm.docx");

	
	    return new ResponseEntity<>(docxContent, headers, HttpStatus.OK);
	}

	
}
