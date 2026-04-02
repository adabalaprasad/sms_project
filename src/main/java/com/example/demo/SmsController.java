package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@CrossOrigin(origins = "*")
@RestController
@RequestMapping("/sms")
public class SmsController 
{
    @Autowired
    private SmsService smsService;

    @Value("${app.download.url}")
    private String downloadUrl;

    @PostMapping("/send")
    public ResponseEntity<String> sendSms(@RequestBody SmsRequest request) 
    {
        try 
        {
        	System.out.println("Controller reached");
            System.out.println("Phone: " + request.getPhoneNumber());
            smsService.sendSms(request.getPhoneNumber(), downloadUrl);

            return ResponseEntity.ok("SMS Sent Successfully");
            
        } 
        catch (Exception e) 
        {
            return ResponseEntity.badRequest().body("SMS Failed: " + e.getMessage());
        }
        
    }
}