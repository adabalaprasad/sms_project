package com.example.demo;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import com.twilio.Twilio;
import com.twilio.rest.api.v2010.account.Message;
import com.twilio.type.PhoneNumber;

@Service
public class SmsService
{

    @Value("${twilio.account.sid}")
    private String accountSid;

    @Value("${twilio.auth.token}")
    private String authToken;

    @Value("${twilio.phone.number}")
    private String twilioPhoneNumber;

    public void sendSms(String toPhoneNumber, String url) {

        Twilio.init(accountSid, authToken);

       long expiryTime = System.currentTimeMillis() + (5L * 24 * 60 * 60 * 1000);
       // long expiryTime = System.currentTimeMillis() + (2 * 60 * 1000);

        String finalUrl = url + "?expiry=" + expiryTime;

        String messageBody =
                "Hello " + toPhoneNumber + ",\n" +
                "Your Student Registration Excel report is ready.\n" +
                "Download here: " + finalUrl + "\n" +
                "Valid for 5 days.\n" +
                "Thank you.";

        Message.creator(
                new PhoneNumber(toPhoneNumber),
                new PhoneNumber(twilioPhoneNumber),
                messageBody
        ).create();

        System.out.println("SMS Sent Successfully");
    }
}