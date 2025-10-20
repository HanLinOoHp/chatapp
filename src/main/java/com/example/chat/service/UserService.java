package com.example.chat.service;

import com.example.chat.dto.ChatSummary;
import com.example.chat.model.User;
import com.example.chat.model.ChatMessageEntity;
import com.example.chat.repository.UserRepository;
import com.example.chat.repository.ChatMessageRepository;
import org.springframework.stereotype.Service;

import java.util.Comparator;
import java.util.List;

@Service
public class UserService {
    private final UserRepository repo;
    private final ChatMessageRepository chatMessageRepository;

    public UserService(UserRepository repo,
            ChatMessageRepository chatMessageRepository) {
        this.repo = repo;
        this.chatMessageRepository = chatMessageRepository;
    }

    public boolean deleteChatBetweenUsers(String user1, String user2) {
        try {
            chatMessageRepository.deleteChatBetweenUsers(user1, user2, user2, user1);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public void blockUser(String blockerUsername, String blockedUsername) {
        if (blockerUsername.equals(blockedUsername))
            return;

        User blocker = repo.findByUsername(blockerUsername)
                .orElseThrow(() -> new RuntimeException("Blocker user not found"));
        User blocked = repo.findByUsername(blockedUsername)
                .orElseThrow(() -> new RuntimeException("Blocked user not found"));

        if (!blocker.getBlockedUsers().contains(blocked)) {
            blocker.getBlockedUsers().add(blocked);
            repo.save(blocker);
        }
    }

    public List<User> getBlockedUsers(String username) {
        User user = repo.findByUsername(username)
                .orElseThrow(() -> new RuntimeException("User not found"));
        return user.getBlockedUsers();
    }

    public void unblockUser(String blockerUsername, String blockedUsername) {
        User blocker = repo.findByUsername(blockerUsername)
                .orElseThrow(() -> new RuntimeException("User not found"));

        blocker.getBlockedUsers().removeIf(u -> u.getUsername().equals(blockedUsername));
        repo.save(blocker);
    }

    public List<ChatSummary> getChatSummaries(String currentUser) {
        User me = repo.findByUsername(currentUser)
                .orElseThrow(() -> new RuntimeException("User not found"));

        // collect blocked usernames (you blocked)
        List<String> blockedUsernames = me.getBlockedUsers().stream()
                .map(User::getUsername)
                .toList();

        // collect users who have blocked you
        List<User> blockedByOthers = repo.findAll().stream()
                .filter(u -> u.getBlockedUsers().stream()
                        .anyMatch(b -> b.getUsername().equals(currentUser)))
                .toList();

        List<String> blockedByUsernames = blockedByOthers.stream()
                .map(User::getUsername)
                .toList();

        // get chat partners (people you have messages with)
        List<String> chatPartners = chatMessageRepository.findChatPartners(currentUser);

        return chatPartners.stream()
                .distinct()
                .filter(username -> !blockedUsernames.contains(username) && // not blocked by you
                        !blockedByUsernames.contains(username)) // not blocked you
                .map(username -> repo.findByUsername(username).orElse(null))
                .filter(u -> u != null)
                .map(user -> {
                    ChatSummary summary = new ChatSummary();
                    summary.setUsername(user.getUsername());
                    summary.setProfilePic(user.getProfilePic());

                    ChatMessageEntity lastMsg = chatMessageRepository
                            .findTopBySenderAndReceiverOrSenderAndReceiverOrderByTimestampDesc(
                                    currentUser, user.getUsername(),
                                    user.getUsername(), currentUser);

                    if (lastMsg != null) {
                        summary.setLastMessage(lastMsg.getContent());
                        summary.setLastTime(lastMsg.getTimestamp());
                    } else {
                        summary.setLastMessage("");
                        summary.setLastTime(null);
                    }

                    return summary;
                })
                .sorted(Comparator.comparing(ChatSummary::getLastTime,
                        Comparator.nullsLast(Comparator.reverseOrder())))
                .toList();
    }

}
