package com.example.chat.repository;

import com.example.chat.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import java.util.Optional;

import java.util.List;

public interface UserRepository extends JpaRepository<User, Long> {
    boolean existsByUsername(String username);

    Optional<User> findByUsername(String username);

    // For listing all users except the current one
    List<User> findByUsernameNot(String username);

    Optional<User> findByEmail(String email);

    Optional<User> findByPhoneNo(String phoneNo);

    boolean existsByEmail(String email);

    boolean existsByPhoneNo(String phoneNo);

    List<User> findByUsernameContainingIgnoreCase(String username);

}
