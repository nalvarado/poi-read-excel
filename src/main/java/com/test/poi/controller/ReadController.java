package com.test.poi.controller;

import java.util.List;
import java.util.Map;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.test.poi.util.ExcelUtil;

import lombok.extern.slf4j.Slf4j;

@RestController
@RequestMapping("/read")
@Slf4j
public class ReadController {

	@GetMapping(value = "/v1")
	public ResponseEntity<String> readV1() throws Exception {
		log.info("Iniciamos la lectura de archivo");
		List<Object> returnList = ExcelUtil.readFolder("C:\\ExcelTest");
		for (int i = 0; i < returnList.size(); i++) {
			List<Map<String, String>> maps = (List<Map<String, String>>) returnList.get(i);
			for (int j = 0; j < maps.size(); j++) {
				System.out.println(maps.get(j).toString());
			}
			System.out.println("-------------------- Línea de corte de lista manual ------------------- ---- ");
		}
		return ResponseEntity.ok("Hello World!");

	}

	@PostMapping(value = "/v2")
	public ResponseEntity<String> readV2(@RequestParam("file") MultipartFile file) throws Exception {
		log.info("Iniciamos la lectura de archivo");

		log.info(file.getOriginalFilename());

		List<Map<String, String>> maps = (List<Map<String, String>>) ExcelUtil.readExcelMultipart(file);

		for (int j = 0; j < maps.size(); j++) {
			System.out.println(maps.get(j).toString());
		}
		System.out.println("-------------------- Línea de corte de lista manual ------------------- ---- ");

		return ResponseEntity.ok("Hello World!");

	}

}
