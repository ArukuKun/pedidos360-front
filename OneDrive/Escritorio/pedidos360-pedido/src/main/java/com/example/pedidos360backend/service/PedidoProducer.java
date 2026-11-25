package com.example.pedidos360backend.service;

import com.example.pedidos360backend.config.RabbitMQConfig;
import com.example.pedidos360backend.dto.EmailDTO;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.amqp.rabbit.core.RabbitTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.kafka.annotation.KafkaListener;
import org.springframework.kafka.core.KafkaTemplate;
import org.springframework.stereotype.Service;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

@Service
public class PedidoProducer {

    @Autowired
    private RabbitTemplate rabbitTemplate;

    @Autowired
    private KafkaTemplate<String, String> kafkaTemplate;

    private final Map<String, Integer> puntajesRecomendacion = new ConcurrentHashMap<>();
    private final Set<String> tarjetasBloqueadas = ConcurrentHashMap.newKeySet();
    private final ObjectMapper objectMapper = new ObjectMapper();

    public void notificarCreacionPedido(String correoDestino, Long idPedido, Double total) {
        String asunto = "Confirmación de Pedido #" + idPedido + " - Pedidos360";
        String cuerpo = "¡Hola!\n\nTu pedido #" + idPedido + " por un total de $" + total + " ha sido procesado exitosamente.\n\n¡Gracias por tu compra!";

        EmailDTO emailDTO = new EmailDTO(correoDestino, asunto, cuerpo);

        rabbitTemplate.convertAndSend(RabbitMQConfig.QUEUE_NAME, emailDTO);
        System.out.println(">>> [RABBITMQ PRODUCTOR] Evento de pedido #" + idPedido + " enviado a 'emailQueue'");
    }

    public void notificarDescuentoInventario(Long idPedido, Long productoId, Integer cantidad) {
        String mensajeJson = String.format(
            "{\"idPedido\":%d,\"productoId\":%d,\"cantidad\":%d}",
            idPedido, productoId, cantidad
        );
        kafkaTemplate.send("pedidos-inventario", mensajeJson);
        System.out.println(">>> [KAFKA PRODUCTOR] Descuento enviado a tópico 'pedidos-inventario' para pedido #" + idPedido);
    }

    public void notificarPago(Long idPedido, Double total) {
        Map<String, Object> mensaje = Map.of(
            "idPedido", idPedido,
            "total", total,
            "estado", "PAGADO"
        );
        rabbitTemplate.convertAndSend(RabbitMQConfig.COLA_PAGOS, mensaje);
        System.out.println(">>> [RABBITMQ PRODUCTOR] Pago enviado a 'cola_pagos' para pedido #" + idPedido);
    }

    public void notificarAlertaFraude(String tarjeta, String cliente, Long idPedido, String tipoAlerta, String detalle) {
        String tarjetaNormalizada = tarjeta != null && !tarjeta.isBlank() ? tarjeta.trim() : "4532";
        tarjetasBloqueadas.add(tarjetaNormalizada);

        String mensajeJson = String.format(
            "{\"tarjeta\":\"%s\",\"cliente\":\"%s\",\"idPedido\":%d,\"tipoAlerta\":\"%s\",\"detalle\":\"%s\",\"bloqueadoPermanente\":true,\"timestamp\":%d}",
            tarjetaNormalizada.replace("\"", "'"),
            cliente != null ? cliente.replace("\"", "'") : "desconocido",
            idPedido,
            tipoAlerta,
            detalle.replace("\"", "'"),
            System.currentTimeMillis()
        );
        kafkaTemplate.send("alertas-fraude", tarjetaNormalizada, mensajeJson);
        System.out.println(">>> [KAFKA ANTIFRAUDE] Tarjeta " + tarjetaNormalizada + " bloqueada permanentemente: " + mensajeJson);
    }

    public boolean estaTarjetaBloqueada(String tarjeta) {
        if (tarjeta == null || tarjeta.isBlank()) return false;
        return tarjetasBloqueadas.contains(tarjeta.trim());
    }

    public void desbloquearTarjeta(String tarjeta) {
        if (tarjeta != null && !tarjeta.isBlank()) {
            tarjetasBloqueadas.remove(tarjeta.trim());
            System.out.println(">>> [ANTIFRAUDE] Bloqueo removido para tarjeta: " + tarjeta);
        }
    }

    public void registrarInteraccionProducto(String usuario, String producto, String accion, int peso) {
        String mensajeJson = String.format(
            "{\"usuario\":\"%s\",\"producto\":\"%s\",\"accion\":\"%s\",\"peso\":%d,\"timestamp\":%d}",
            usuario.replace("\"", "'"),
            producto.replace("\"", "'"),
            accion,
            peso,
            System.currentTimeMillis()
        );
        kafkaTemplate.send("interacciones-productos", producto, mensajeJson);
        System.out.println(">>> [KAFKA INTERACCION] " + accion + " (peso " + peso + ") -> " + producto);
    }

    @KafkaListener(topics = "interacciones-productos", groupId = "pedidos-recomendaciones-group")
    public void consumirInteraccionProducto(String mensajeJson) {
        try {
            JsonNode node = objectMapper.readTree(mensajeJson);
            String producto = node.path("producto").asText("");
            int peso = node.path("peso").asInt(0);
            if (!producto.isEmpty() && peso > 0) {
                puntajesRecomendacion.merge(producto, peso, Integer::sum);
                System.out.println(">>> [KAFKA RECOMENDADOR] Puntaje actualizado: " + producto + " = " + puntajesRecomendacion.get(producto));
            }
        } catch (Exception e) {
            System.err.println("Error procesando interacción en Kafka: " + e.getMessage());
        }
    }

    public Map<String, Integer> obtenerProductosRecomendados() {
        return puntajesRecomendacion.entrySet().stream()
            .sorted(Map.Entry.<String, Integer>comparingByValue().reversed())
            .collect(Collectors.toMap(
                Map.Entry::getKey,
                Map.Entry::getValue,
                (e1, e2) -> e1,
                LinkedHashMap::new
            ));
    }
}