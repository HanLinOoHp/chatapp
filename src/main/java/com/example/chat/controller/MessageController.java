package com.example.chat.controller;

import com.example.chat.dto.ChatMessage;
import com.example.chat.model.ChatMessageEntity;
import com.example.chat.repository.ChatMessageRepository;

import org.springframework.messaging.handler.annotation.MessageMapping;
import org.springframework.messaging.handler.annotation.SendTo;
import org.springframework.stereotype.Controller;

import com.example.chat.dto.ChatMessage;
import org.springframework.messaging.handler.annotation.MessageMapping;
import org.springframework.stereotype.Controller;
import org.springframework.messaging.simp.SimpMessagingTemplate;

import java.security.Principal;

@Controller
public class MessageController {

    private final SimpMessagingTemplate messagingTemplate;
    private final ChatMessageRepository chatMessageRepository;

    public MessageController(SimpMessagingTemplate messagingTemplate,
            ChatMessageRepository chatMessageRepository) {
        this.messagingTemplate = messagingTemplate;
        this.chatMessageRepository = chatMessageRepository;
    }

    @MessageMapping("/chat.private")
    public void sendPrivate(ChatMessage message, Principal principal) {
        message.setFrom(principal.getName());
        ChatMessageEntity entity = new ChatMessageEntity();
        entity.setSender(message.getFrom());
        entity.setReceiver(message.getTo());
        entity.setContent(message.getText());
        chatMessageRepository.save(entity);

        messagingTemplate.convertAndSendToUser(
                message.getTo(),
                "/queue/messages",
                message);
    }

}
