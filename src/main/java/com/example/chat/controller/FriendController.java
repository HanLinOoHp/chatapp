package com.example.chat.controller;

import com.example.chat.model.User;
import com.example.chat.repository.UserRepository;
import com.example.chat.service.FriendService;
import com.example.chat.service.UserService;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.security.Principal;

@Controller
public class FriendController {

    private final FriendService friendService;
    private final UserRepository userRepository;
    private final UserService userService;

    public FriendController(FriendService friendService, UserRepository userRepository, UserService userService) {
        this.friendService = friendService;
        this.userRepository = userRepository;
        this.userService = userService;
    }

    // Show add friend page
    @GetMapping("/addfriend")
    public String addFriendPage() {
        return "addfriend";
    }

    // Handle friend adding
    @PostMapping("/addfriend")
    public String addFriend(@RequestParam("contact") String contact, Model model, Principal principal) {
        String username = principal.getName();
        User currentUser = userRepository.findByUsername(username)
                .orElseThrow(() -> new RuntimeException("User not found"));

        String message = friendService.addFriend(currentUser, contact);
        model.addAttribute("message", message);

        return "addfriend";
    }

    @GetMapping("/users")
    public String usersPage(Model model, Principal principal) {
        String username = principal.getName();
        User currentUser = userRepository.findByUsername(username)
                .orElseThrow(() -> new RuntimeException("User not found"));

        // This now returns List<ChatSummary>
        model.addAttribute("users", friendService.getFriends(currentUser));
        return "users";
    }

    @GetMapping("/friends")
    public String friendPage(Model model, Principal principal) {
        String username = principal.getName();
        User currentUser = userRepository.findByUsername(username)
                .orElseThrow(() -> new RuntimeException("User not found"));

        // This now returns List<ChatSummary>
        model.addAttribute("users", friendService.getFriends(currentUser));
        return "friends";
    }

    @DeleteMapping("/userinfo/delete/{username}")
    @ResponseBody
    public ResponseEntity<?> deleteChat(@PathVariable String username, Principal principal) {
        String currentUser = principal.getName();
        boolean success = userService.deleteChatBetweenUsers(currentUser, username);
        if (success) {
            return ResponseEntity.ok().build();
        } else {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to delete chat.");
        }
    }

    @GetMapping("/userinfo/{receiver}")
    public String showUserInfo(@PathVariable String receiver, Model model) {
        User user = userRepository.findByUsername(receiver)
                .orElse(null);
        if (user == null) {
            return "redirect:/users?notfound";
        }
        model.addAttribute("user", user);
        model.addAttribute("receiver", receiver);
        return "userinfo";
    }

}
