package com.example.chat.controller;

import com.example.chat.dto.ChatSummary;
import com.example.chat.dto.UserDTO;
import com.example.chat.model.User;
import com.example.chat.repository.UserRepository;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import com.example.chat.service.FriendService;
import com.example.chat.service.UserService;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.security.Principal;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/users")

public class UserController {

    private final UserRepository userRepository;
    private final FriendService friendService;

    public UserController(UserRepository userRepository, FriendService friendService) {
        this.userRepository = userRepository;
        this.friendService = friendService;
    }

    @GetMapping("/friends")
    public List<ChatSummary> getFriends(Principal principal) {
        String currentUsername = principal.getName();
        User currentUser = userRepository.findByUsername(currentUsername)
                .orElseThrow(() -> new RuntimeException("User not found"));

        return friendService.getFriends(currentUser);
    }
}
