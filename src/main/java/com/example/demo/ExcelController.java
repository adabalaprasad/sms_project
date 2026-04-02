package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
public class ExcelController
{

    @Autowired
    private ExcelService service;

    @GetMapping("/download")
    public ResponseEntity<?> downloadExcel(@RequestParam long expiry) throws Exception {

        if (System.currentTimeMillis() > expiry) {
            return ResponseEntity.badRequest().body("Link Expired");
        }

        byte[] data = service.exportExcel();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=students.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(data);
    }
}