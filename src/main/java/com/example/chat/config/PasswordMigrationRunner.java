package com.example.chat.config;

import com.example.chat.model.User;
import com.example.chat.repository.UserRepository;
import org.springframework.boot.CommandLineRunner;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Component;

@Component
public class PasswordMigrationRunner implements CommandLineRunner {

    private final UserRepository repo;
    private final PasswordEncoder encoder;

    public PasswordMigrationRunner(UserRepository repo, PasswordEncoder encoder) {
        this.repo = repo;
        this.encoder = encoder;
    }

    @Override
    public void run(String... args) throws Exception {
        for (User u : repo.findAll()) {
            String pw = u.getPassword();
            if (pw == null)
                continue;
            // simple heuristic: BCrypt hashes start with $2a$ or $2b$ or $2y$
            if (!(pw.startsWith("$2a$") || pw.startsWith("$2b$") || pw.startsWith("$2y$"))) {
                u.setPassword(encoder.encode(pw));
                repo.save(u);
            }
        }
    }
}
