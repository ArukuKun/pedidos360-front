package com.example.pedidos360backend.dto;

import java.io.Serializable;
import com.fasterxml.jackson.annotation.JsonProperty;

public class EmailDTO implements Serializable {

    @JsonProperty("destinatario")
    private String to;

    private String subject;

    @JsonProperty("mensaje")
    private String body;

    public EmailDTO() {
    }

    public EmailDTO(String to, String subject, String body) {
        this.to = to;
        this.subject = subject;
        this.body = body;
    }

    public String getTo() {
        return to;
    }

    public void setTo(String to) {
        this.to = to;
    }

    public String getSubject() {
        return subject;
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public String getBody() {
        return body;
    }

    public void setBody(String body) {
        this.body = body;
    }
}