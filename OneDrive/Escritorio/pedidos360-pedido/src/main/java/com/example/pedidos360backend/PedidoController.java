package com.example.pedidos360backend;

import com.example.pedidos360backend.service.PedidoProducer;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.ArrayList;
import java.util.Deque;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentLinkedDeque;

@RestController
@RequestMapping("/v1/pedidos")
public class PedidoController {

    @Autowired
    private PedidoProducer pedidoProducer;

    private final List<String> pedidos = new ArrayList<>();
    private final Deque<Long> tiemposClicksPago = new ConcurrentLinkedDeque<>();
    private final ObjectMapper objectMapper = new ObjectMapper();

    private static class ItemPedido {
        final String nombre;
        final int cantidad;

        ItemPedido(String nombre, int cantidad) {
            this.nombre = nombre;
            this.cantidad = cantidad;
        }
    }

    @GetMapping
    public List<String> obtenerPedidos() {
        return pedidos;
    }

    @GetMapping("/recomendados")
    public Map<String, Integer> obtenerRecomendados() {
        return pedidoProducer.obtenerProductosRecomendados();
    }

    // Consultado por Webpay para saber si la tarjeta específica está bloqueada
    @GetMapping("/tarjeta-bloqueada/{tarjeta}")
    public Map<String, Object> verificarTarjetaBloqueada(@PathVariable String tarjeta) {
        boolean bloqueada = pedidoProducer.estaTarjetaBloqueada(tarjeta);
        return Map.of(
            "tarjeta", tarjeta,
            "bloqueada", bloqueada,
            "mensaje", bloqueada
                ? "La tarjeta •••• " + tarjeta + " está bloqueada permanentemente. Cambia de tarjeta para pagar."
                : "Tarjeta habilitada para operar."
        );
    }

    // Endpoint para limpiar el bloqueo de una tarjeta en pruebas
    @DeleteMapping("/tarjeta-bloqueada/{tarjeta}")
    public Map<String, Object> desbloquearTarjeta(@PathVariable String tarjeta) {
        pedidoProducer.desbloquearTarjeta(tarjeta);
        tiemposClicksPago.clear();
        return Map.of(
            "tarjeta", tarjeta,
            "bloqueada", false,
            "mensaje", "Bloqueo removido para tarjeta •••• " + tarjeta
        );
    }

    @PostMapping("/interaccion")
    public Map<String, Object> registrarInteraccion(@RequestBody Map<String, Object> body) {
        String usuario = String.valueOf(body.getOrDefault("usuario", "anonimo"));
        String producto = String.valueOf(body.getOrDefault("producto", ""));
        String accion = String.valueOf(body.getOrDefault("accion", "VER")).toUpperCase();

        int peso = 1; // VER
        if ("COMPRAR".equals(accion)) {
            peso = 3;
        } else if ("GUARDAR".equals(accion)) {
            peso = 2;
        }

        if (producto != null && !producto.trim().isEmpty()) {
            pedidoProducer.registrarInteraccionProducto(usuario, producto, accion, peso);
        }

        return Map.of(
            "mensaje", "Interacción registrada en Kafka",
            "producto", producto,
            "accion", accion,
            "peso", peso,
            "recomendados", pedidoProducer.obtenerProductosRecomendados()
        );
    }

    @PostMapping
    public Map<String, Object> crearPedido(@RequestBody String pedidoRaw) {
        List<ItemPedido> itemsProcesados = new ArrayList<>();
        double montoTotal = 100.0;
        String cliente = "mb.espinoza@duocuc.cl";
        String tarjeta = "4532";
        String tipoEvento = "CREAR_PEDIDO";
        boolean forzarFraude = false;
        int numeroIntento = 1;

        try {
            JsonNode root = objectMapper.readTree(pedidoRaw);
            if (root.has("montoTotal")) {
                montoTotal = root.path("montoTotal").asDouble(100.0);
            }
            if (root.has("cliente")) {
                cliente = root.path("cliente").asText("mb.espinoza@duocuc.cl");
            }
            if (root.has("tarjeta")) {
                tarjeta = root.path("tarjeta").asText("4532");
            }
            if (root.has("tipoEvento")) {
                tipoEvento = root.path("tipoEvento").asText("CREAR_PEDIDO");
            }
            if (root.has("forzarFraude")) {
                forzarFraude = root.path("forzarFraude").asBoolean(false);
            }
            if (root.has("numeroIntento")) {
                numeroIntento = root.path("numeroIntento").asInt(1);
            }
            if (root.has("items") && root.get("items").isArray()) {
                for (JsonNode itemNode : root.get("items")) {
                    String nombre = itemNode.path("nombre").asText("Producto");
                    int cantidad = Math.max(1, itemNode.path("cantidad").asInt(1));
                    itemsProcesados.add(new ItemPedido(nombre, cantidad));
                }
            }
        } catch (Exception ignored) {
        }

        if (itemsProcesados.isEmpty()) {
            itemsProcesados.add(new ItemPedido("Producto", 1));
        }

        // CASO A: Intento de pago -> bloquea SOLO la tarjeta usada si hay ráfaga (>3 clicks en <5s)
        if ("INTENTO_PAGO".equalsIgnoreCase(tipoEvento)) {
            long ahora = System.currentTimeMillis();
            synchronized (tiemposClicksPago) {
                while (!tiemposClicksPago.isEmpty() && (ahora - tiemposClicksPago.peekFirst() > 5000L)) {
                    tiemposClicksPago.pollFirst();
                }
                tiemposClicksPago.addLast(ahora);
            }

            int clicksEnVentana5s = tiemposClicksPago.size();
            boolean tarjetaYaBloqueada = pedidoProducer.estaTarjetaBloqueada(tarjeta);
            boolean esFraude = tarjetaYaBloqueada || clicksEnVentana5s > 3 || forzarFraude;
            List<String> alertasFraude = new ArrayList<>();

            if (esFraude) {
                long idRef = pedidos.size() + numeroIntento;
                for (ItemPedido item : itemsProcesados) {
                    for (int u = 1; u <= item.cantidad; u++) {
                        String detalleIndividual = "Fraude con tarjeta •••• " + tarjeta + " al comprar " + item.nombre
                            + " (Click #" + numeroIntento + ") — Tarjeta bloqueada permanentemente";
                        alertasFraude.add(detalleIndividual);
                        pedidoProducer.notificarAlertaFraude(tarjeta, cliente, idRef, "RAFAGA_BOTON_PAGAR_5S", detalleIndividual);
                    }
                }
            }

            Map<String, Object> respPago = new LinkedHashMap<>();
            respPago.put("tipoEvento", "INTENTO_PAGO");
            respPago.put("tarjeta", tarjeta);
            respPago.put("numeroIntento", numeroIntento);
            respPago.put("clicksEn5Segundos", clicksEnVentana5s);
            respPago.put("fraude", esFraude);
            respPago.put("tarjetaBloqueada", esFraude);
            respPago.put("alertasFraude", alertasFraude);
            if (esFraude) {
                respPago.put("mensaje", "Tarjeta •••• " + tarjeta + " bloqueada por fraude. Selecciona otra tarjeta para pagar.");
            } else {
                respPago.put("mensaje", "Intento de pago #" + numeroIntento + " con tarjeta •••• " + tarjeta + ": OK");
            }
            return respPago;
        }

        // CASO B: Confirmación de pedido normal desde el carrito
        pedidos.add(pedidoRaw);
        long idPedido = pedidos.size();

        for (ItemPedido item : itemsProcesados) {
            for (int u = 1; u <= item.cantidad; u++) {
                pedidoProducer.registrarInteraccionProducto(cliente, item.nombre, "COMPRAR", 3);
            }
        }

        pedidoProducer.notificarCreacionPedido(cliente, idPedido, montoTotal);
        pedidoProducer.notificarDescuentoInventario(idPedido, 1L, 1);
        pedidoProducer.notificarPago(idPedido, montoTotal);

        Map<String, Object> respuesta = new LinkedHashMap<>();
        respuesta.put("idPedido", idPedido);
        respuesta.put("fraude", false);
        respuesta.put("alertasFraude", List.of());
        respuesta.put("mensaje", "Pedido #" + idPedido + " confirmado y eventos publicados en RabbitMQ y Kafka");
        return respuesta;
    }

    @PutMapping("/{id}")
    public String modificarPedido(@PathVariable int id, @RequestBody String pedidoModificado) {
        if (id >= 0 && id < pedidos.size()) {
            pedidos.set(id, pedidoModificado);
            return "Carrito actualizado";
        }
        return "Error: Pedido no encontrado";
    }

    @DeleteMapping("/{id}")
    public String eliminarPedido(@PathVariable int id) {
        if (id >= 0 && id < pedidos.size()) {
            pedidos.remove(id);
            return "Eliminado del carrito";
        }
        return "Error: Pedido no encontrado";
    }
}