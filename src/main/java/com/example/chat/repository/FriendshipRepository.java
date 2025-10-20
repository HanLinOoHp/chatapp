package com.example.chat.repository;

import com.example.chat.model.Friendship;
import com.example.chat.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;
import java.util.Optional;

public interface FriendshipRepository extends JpaRepository<Friendship, Long> {
    List<Friendship> findByUser(User user);

    Optional<Friendship> findByUserAndFriend(User user, User friend);

    @Query("""
                SELECT f FROM Friendship f
                WHERE (f.user = :user OR f.friend = :user)
            """)
    List<Friendship> findFriendships(User user);

    @Query("SELECT f.friend FROM Friendship f WHERE f.user = :user")
    List<User> findFriendsOfUser(@Param("user") User user);

}
