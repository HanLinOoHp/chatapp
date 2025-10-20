package com.example.chat.repository;

import com.example.chat.model.ChatMessageEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.transaction.annotation.Transactional;
import java.util.List;

public interface ChatMessageRepository extends JpaRepository<ChatMessageEntity, Long> {

        // For ChatSummary (last message preview in users list)
        ChatMessageEntity findTopBySenderAndReceiverOrSenderAndReceiverOrderByTimestampDesc(
                        String sender1, String receiver1, String sender2, String receiver2);

        // For Private Chat (full conversation)
        List<ChatMessageEntity> findBySenderAndReceiverOrSenderAndReceiverOrderByTimestampAsc(
                        String sender1, String receiver1, String sender2, String receiver2);

        // Delete all messages between two users
        @Modifying
        @Transactional
        @Query("DELETE FROM ChatMessageEntity c " +
                        "WHERE (c.sender = :sender1 AND c.receiver = :receiver1) " +
                        "   OR (c.sender = :sender2 AND c.receiver = :receiver2)")
        void deleteChatBetweenUsers(@Param("sender1") String sender1,
                        @Param("receiver1") String receiver1,
                        @Param("sender2") String sender2,
                        @Param("receiver2") String receiver2);

        @Transactional
        void deleteBySenderAndReceiver(String sender, String receiver);

        @Query("SELECT DISTINCT CASE WHEN c.sender = :username THEN c.receiver ELSE c.sender END " +
                        "FROM ChatMessageEntity c WHERE c.sender = :username OR c.receiver = :username")
        List<String> findChatPartners(@Param("username") String username);

}
