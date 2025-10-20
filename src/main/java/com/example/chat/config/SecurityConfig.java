package com.example.chat.config;

import com.example.chat.repository.UserRepository;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.core.userdetails.*;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.config.Customizer;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;

@Configuration
public class SecurityConfig {

        private final UserRepository userRepository;

        public SecurityConfig(UserRepository userRepository) {
                this.userRepository = userRepository;
        }

        // Load users from DB
        @Bean
        public UserDetailsService userDetailsService() {
                return username -> userRepository.findByUsername(username)
                                .map(u -> User.withUsername(u.getUsername())
                                                .password(u.getPassword())
                                                .roles("USER")
                                                .build())
                                .orElseThrow(() -> new UsernameNotFoundException("User not found"));
        }

        // Use BCrypt for password hashing
        @Bean
        public PasswordEncoder passwordEncoder() {
                return new BCryptPasswordEncoder();
        }

        // Security filter chain
        @Bean
        public SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
                http
                                // Authorization rules
                                .authorizeHttpRequests(auth -> auth
                                                .requestMatchers(
                                                                "/",
                                                                "/register",
                                                                "/login",
                                                                "/css/**",
                                                                "/js/**",
                                                                "/ws/**",
                                                                "/uploads/**" // allow profile image access
                                                ).permitAll()
                                                .anyRequest().authenticated())

                                // Login page config
                                .formLogin(form -> form
                                                .loginPage("/login")
                                                .defaultSuccessUrl("/users", true)
                                                .permitAll())

                                // Logout support (proper session cleanup)
                                .logout(logout -> logout
                                                .logoutUrl("/logout") // default logout URL
                                                .logoutSuccessUrl("/login?logout") // redirect after logout
                                                .invalidateHttpSession(true) // clear session
                                                .clearAuthentication(true) // clear authentication
                                                .deleteCookies("JSESSIONID") // remove session cookie
                                                .permitAll())

                                // Enable CSRF (needed for forms)
                                .csrf(Customizer.withDefaults());

                return http.build();
        }
}
