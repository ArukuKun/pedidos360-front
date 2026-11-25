package com.example.pedidos360backend.config;

import org.springframework.amqp.core.Queue;
import org.springframework.amqp.support.converter.Jackson2JsonMessageConverter;
import org.springframework.amqp.support.converter.MessageConverter;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class RabbitMQConfig {

    public static final String QUEUE_NAME = "emailQueue";
    public static final String COLA_INVENTARIO = "cola_inventario";
    public static final String COLA_PAGOS = "cola_pagos";

    @Bean
    public Queue emailQueue() {
        return new Queue(QUEUE_NAME, true);
    }

    @Bean
    public Queue colaInventario() {
        return new Queue(COLA_INVENTARIO, true);
    }

    @Bean
    public Queue colaPagos() {
        return new Queue(COLA_PAGOS, true);
    }

    @Bean
    public MessageConverter jsonMessageConverter() {
        return new Jackson2JsonMessageConverter();
    }
}