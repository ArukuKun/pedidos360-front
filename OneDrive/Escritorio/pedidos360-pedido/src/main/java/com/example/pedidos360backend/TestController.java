package com.example.pedidos360backend;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api")
public class TestController {

    @GetMapping("/test")
    public String probarSeguridad() {
        return "¡Conexión exitosa! Angular y Spring Boot por fin están hablando.";
    }
}