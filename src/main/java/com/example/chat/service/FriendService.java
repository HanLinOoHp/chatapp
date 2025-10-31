package com.example.chat.service;

import com.example.chat.dto.ChatSummary;
import com.example.chat.model.ChatMessageEntity;
import com.example.chat.model.Friendship;
import com.example.chat.model.User;
import com.example.chat.repository.ChatMessageRepository;
import com.example.chat.repository.FriendshipRepository;
import com.example.chat.repository.UserRepository;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

@Service
public class FriendService {
    private final UserRepository userRepository;
    private final ChatMessageRepository chatMessageRepository;
    private final FriendshipRepository friendshipRepository;

    public FriendService(UserRepository userRepository,
            ChatMessageRepository chatMessageRepository,
            FriendshipRepository friendshipRepository) {
        this.userRepository = userRepository;
        this.chatMessageRepository = chatMessageRepository;
        this.friendshipRepository = friendshipRepository;
    }

    public List<ChatSummary> getFriends(User currentUser) {

        // Users you blocked
        List<String> blockedUsernames = currentUser.getBlockedUsers()
                .stream()
                .map(User::getUsername)
                .toList();

        // Users who blocked you
        List<String> blockedByOthers = userRepository.findAll().stream()
                .filter(u -> u.getBlockedUsers().stream()
                        .anyMatch(b -> b.getUsername().equals(currentUser.getUsername())))
                .map(User::getUsername)
                .toList();

        // Get friends from friendship table
        List<User> friends = friendshipRepository.findFriendsOfUser(currentUser);

        // Filter blocked users
        return friends.stream()
                .filter(f -> !blockedUsernames.contains(f.getUsername()) &&
                        !blockedByOthers.contains(f.getUsername()))
                .map(f -> {
                    ChatSummary summary = new ChatSummary();
                    summary.setUsername(f.getUsername());
                    summary.setProfilePic(f.getProfilePic());

                    // Optional: get last chat message
                    ChatMessageEntity lastMsg = chatMessageRepository
                            .findTopBySenderAndReceiverOrSenderAndReceiverOrderByTimestampDesc(
                                    currentUser.getUsername(), f.getUsername(),
                                    f.getUsername(), currentUser.getUsername());

                    if (lastMsg != null) {
                        summary.setLastMessage(lastMsg.getContent());
                        summary.setLastTime(lastMsg.getTimestamp());
                    } else {
                        summary.setLastMessage("");
                        summary.setLastTime(null);
                    }

                    return summary;
                })
                .sorted((c1, c2) -> {
                    if (c1.getLastTime() == null && c2.getLastTime() == null)
                        return 0;
                    if (c1.getLastTime() == null)
                        return 1;
                    if (c2.getLastTime() == null)
                        return -1;
                    return c2.getLastTime().compareTo(c1.getLastTime()); // descending
                })
                .collect(Collectors.toList());
    }

    // Add friend by email or phone
    public String addFriend(User currentUser, String contact) {
        Optional<User> friendOpt;

        if (contact.contains("@")) {
            friendOpt = userRepository.findByEmail(contact);
        } else {
            friendOpt = userRepository.findByPhoneNo(contact);
        }

        if (friendOpt.isEmpty()) {
            return "User not found with that email or phone number.";
        }

        User friend = friendOpt.get();

        // Prevent adding yourself
        if (friend.getId().equals(currentUser.getId())) {
            return "You cannot add yourself.";
        }

        // Prevent duplicate friendship
        if (friendshipRepository.findByUserAndFriend(currentUser, friend).isPresent()
                || friendshipRepository.findByUserAndFriend(friend, currentUser).isPresent()) {
            return "This user is already your friend.";
        }

        // Save two-way friendship
        Friendship friendship1 = new Friendship();
        friendship1.setUser(currentUser);
        friendship1.setFriend(friend);

        Friendship friendship2 = new Friendship();
        friendship2.setUser(friend);
        friendship2.setFriend(currentUser);

        friendshipRepository.save(friendship1);
        friendshipRepository.save(friendship2);

        return "Friend added successfully.";

    }

}
