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
    
    @Autowired
    private StudentRepo repo;

    @GetMapping("/download")
    public ResponseEntity<?> downloadExcel(@RequestParam long expiry) throws Exception 
    {

        if (System.currentTimeMillis() > expiry) 
        {
            return ResponseEntity.badRequest().body("Link Expired");
        }

        byte[] data = service.exportExcel();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=students.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(data);
    }
    
    @GetMapping("/addTest")
    public String addTest() 
    {

        Student s = new Student();

        s.setName("prasad");
        s.setEmail("prasad" + System.currentTimeMillis() + "@test.com");
        s.setPhoneNumber("9999999999");
        s.setPassword("1234");
        s.setRole("USER");
        s.setIsVerified(true);
        s.setOtp("1111");

        repo.save(s);

        return "Inserted";
    }
}