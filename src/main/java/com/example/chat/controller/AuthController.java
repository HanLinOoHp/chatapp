package com.example.chat.controller;

import com.example.chat.model.User;
import com.example.chat.repository.UserRepository;
import com.example.chat.service.UserService;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.*;
import java.security.Principal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Controller
public class AuthController {

    private final UserService userService;
    private final UserRepository userRepository;
    private final PasswordEncoder passwordEncoder;

    public AuthController(UserService userService,
            UserRepository userRepository,
            PasswordEncoder passwordEncoder) {
        this.userService = userService;
        this.userRepository = userRepository;
        this.passwordEncoder = passwordEncoder;
    }

    @GetMapping("/")
    public String root(Principal principal) {
        if (principal != null) {
            return "redirect:/users";
        }
        return "redirect:/login";
    }

    @GetMapping("/login")
    public String showLogin() {
        return "login";
    }

    @GetMapping("/setting")
    public String showSettings(Model model, Principal principal) {
        if (principal == null) {
            return "redirect:/login";
        }

        String username = principal.getName();
        User user = userRepository.findByUsername(username)
                .orElse(null);

        if (user == null) {
            return "redirect:/login";
        }

        model.addAttribute("user", user);
        return "setting";
    }

    @GetMapping("/register")
    public String showRegister() {
        return "register";
    }

    @GetMapping("/aboutus")
    public String aboutPage() {
        return "aboutus";
    }

    @GetMapping("/help")
    public String helpPage() {
        return "help";
    }

    @GetMapping("/blocked")
    public String blockedList(Model model, Principal principal) {
        List<User> blockedUsers = userService.getBlockedUsers(principal.getName());
        model.addAttribute("blockedUsers", blockedUsers);
        return "blocked";
    }

    @PostMapping("/block")
    public String blockUser(@RequestParam("blockUsername") String blockUsername, Principal principal) {
        String currentUser = principal.getName();
        userService.blockUser(currentUser, blockUsername);
        return "redirect:/blocked";
    }

    @PostMapping("/unblock/{username}")
    public String unblockUser(@PathVariable String username, Principal principal) {
        userService.unblockUser(principal.getName(), username);
        return "redirect:/blocked";
    }

    @PostMapping("/register")
    public String doRegister(@RequestParam String username,
            @RequestParam String password,
            @RequestParam String dob,
            @RequestParam String phoneNo,
            @RequestParam String email,
            @RequestParam("profilePic") MultipartFile profilePic,
            Model model) {

        if (userRepository.findByUsername(username).isPresent()) {
            model.addAttribute("error", "Username already taken");
            return "register";
        }
        if (userRepository.findByEmail(email).isPresent()) {
            model.addAttribute("error", "Email already registered");
            return "register";
        }
        if (userRepository.findByPhoneNo(phoneNo).isPresent()) {
            model.addAttribute("error", "Phone number already registered");
            return "register";
        }

        String fileName = null;
        if (profilePic != null && !profilePic.isEmpty()) {
            try {
                fileName = System.currentTimeMillis() + "_" + profilePic.getOriginalFilename();
                String uploadDir = System.getProperty("user.home") + "/chatapp/uploads/profile/";
                Path uploadPath = Paths.get(uploadDir);
                Files.createDirectories(uploadPath);
                Files.copy(profilePic.getInputStream(), uploadPath.resolve(fileName),
                        StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException e) {
                model.addAttribute("error", "Could not upload profile picture.");
                return "register";
            }
        }

        User user = new User();
        user.setUsername(username);
        user.setPassword(passwordEncoder.encode(password));
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
            user.setDob(LocalDate.parse(dob, formatter));
        } catch (Exception e) {
            model.addAttribute("error", "Invalid date format. Use yyyy-MM-dd.");
            return "register";
        }
        user.setPhoneNo(phoneNo);
        user.setEmail(email);
        user.setProfilePic(fileName);

        userRepository.save(user);

        return "redirect:/login?registered";
    }

    @GetMapping("/usersettings")
    public String showUserSettings(Model model, Principal principal) {
        if (principal == null) {
            return "redirect:/login";
        }

        String username = principal.getName();
        User user = userRepository.findByUsername(username)
                .orElse(null);

        if (user == null) {
            return "redirect:/login";
        }

        model.addAttribute("user", user);
        return "usersettings";
    }

    @GetMapping("/personalsetting")
    public String personalsetting(Model model, Principal principal) {
        if (principal == null) {
            return "redirect:/login";
        }

        // Get the logged-in user
        User user = userRepository.findByUsername(principal.getName())
                .orElse(null);

        if (user == null) {
            return "redirect:/login";
        }

        // Add user to the model for Thymeleaf
        model.addAttribute("user", user);

        return "personalsetting";
    }

    @PostMapping("/updateProfile")
    public String updateProfile(@RequestParam String username,
            @RequestParam String email,
            @RequestParam String phoneNo,
            @RequestParam String dob,
            @RequestParam(value = "profilePic", required = false) MultipartFile profilePic,
            Principal principal) throws IOException {

        User currentUser = userRepository.findByUsername(principal.getName())
                .orElseThrow(() -> new RuntimeException("User not found"));

        currentUser.setUsername(username);
        currentUser.setEmail(email);
        currentUser.setPhoneNo(phoneNo);
        try {
            currentUser.setDob(LocalDate.parse(dob));
        } catch (Exception e) {
            // handle invalid date format
        }

        if (profilePic != null && !profilePic.isEmpty()) {
            String fileName = System.currentTimeMillis() + "_" + profilePic.getOriginalFilename();
            String uploadDir = System.getProperty("user.home") + "/chatapp/uploads/profile/";
            Path uploadPath = Paths.get(uploadDir);
            if (!Files.exists(uploadPath)) {
                Files.createDirectories(uploadPath);
            }
            Files.copy(profilePic.getInputStream(), uploadPath.resolve(fileName),
                    StandardCopyOption.REPLACE_EXISTING);
            currentUser.setProfilePic(fileName);
        }

        userRepository.save(currentUser);

        return "redirect:/personalsetting";
    }

}
