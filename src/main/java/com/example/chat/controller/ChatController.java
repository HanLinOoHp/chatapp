package com.example.chat.controller;

import com.example.chat.model.ChatMessageEntity;
import com.example.chat.repository.ChatMessageRepository;
import com.example.chat.repository.UserRepository;
import com.example.chat.service.UserService;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.transaction.annotation.Transactional;
import com.example.chat.service.FriendService;

import java.security.Principal;
import java.util.List;

import com.example.chat.model.User;

@Controller
public class ChatController {
    private final UserService userService;
    private final ChatMessageRepository chatMessageRepository;
    private final UserRepository userRepository;
    private final FriendService friendService;

    public ChatController(UserService userService, ChatMessageRepository chatMessageRepository,
            UserRepository userRepository, FriendService friendService) {
        this.userService = userService;
        this.chatMessageRepository = chatMessageRepository;
        this.userRepository = userRepository;
        this.friendService = friendService;
    }

    @GetMapping("/chat")
    public String chatPageRedirect() {
        return "redirect:/users";
    }

    @GetMapping("/chat/{username}")
    public String chatWith(@PathVariable String username, Model model, Principal principal) {
        String currentUser = principal.getName();

        List<ChatMessageEntity> messages = chatMessageRepository
                .findBySenderAndReceiverOrSenderAndReceiverOrderByTimestampAsc(
                        currentUser, username, username, currentUser);
        User receiver = userRepository.findByUsername(username)
                .orElseThrow(() -> new RuntimeException("Receiver not found"));

        model.addAttribute("receiver", receiver.getUsername());
        model.addAttribute("receiverImage", receiver.getProfilePic());
        model.addAttribute("sender", currentUser);
        model.addAttribute("messages", messages);
        return "privatechat";
    }
}
