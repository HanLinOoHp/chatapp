package com.example.chat.dto;

import lombok.Data;
import java.time.LocalDateTime;

@Data
public class ChatSummary {
    private String username;
    private String lastMessage;
    private LocalDateTime lastTime;
    private String profilePic;
}