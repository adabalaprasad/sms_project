package com.example.demo;


import java.util.concurrent.TimeUnit;
import org.springframework.cache.CacheManager;
import org.springframework.cache.annotation.EnableCaching;
import org.springframework.cache.caffeine.CaffeineCacheManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import com.github.benmanes.caffeine.cache.Caffeine;

@Configuration
@EnableCaching
public class CacheConfig 
{

    @Bean
    public CacheManager cacheManager() 
    {
        CaffeineCacheManager cacheManager = new CaffeineCacheManager("studentExcel");

        cacheManager.setCaffeine(
                Caffeine.newBuilder()
                       // .expireAfterWrite(5, TimeUnit.DAYS)
                .expireAfterWrite(2, TimeUnit.MINUTES)
                        .maximumSize(10)
        );

        return cacheManager;
    }
}
