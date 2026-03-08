import sqlite3
import time

class WeChatDatabase:
    def __init__(self, db_path="wechat_memory.db"):
        self.db_path = db_path
        self._create_table()
    
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, check_same_thread=False)
        # 注册汉明距离函数，直接在 SQL 层完成纠错匹配
        def hamming_distance(h1, h2):
            if not h1 or not h2 or len(h1) != len(h2): return 999
            try:
                return bin(int(h1, 16) ^ int(h2, 16)).count('1')
            except ValueError: return 999
        conn.create_function("hamming", 2, hamming_distance)
        return conn
    
    def _create_table(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                avatar_hash TEXT NOT NULL,
                name_img_hash TEXT NOT NULL,
                nickname TEXT NOT NULL UNIQUE
            )""")
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contact_id INTEGER NOT NULL,
                contact_name TEXT NOT NULL,
                sender TEXT NOT NULL,
                content TEXT NOT NULL,
                timestamp INTEGER NOT NULL,
                FOREIGN KEY (contact_id) REFERENCES contacts (id)
            )""")
            conn.commit()
        finally:
            conn.close()

    def resolve_contact(self, avatar_hash, name_img_hash, raw_nickname):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, nickname FROM contacts 
                WHERE hamming(avatar_hash, ?) <= 5 AND hamming(name_img_hash, ?) <= 5
            """, (avatar_hash, name_img_hash))
            
            row = cursor.fetchone()
            if row: return row[0], row[1]
                
            cursor.execute("SELECT id FROM contacts WHERE nickname = ?", (raw_nickname,))
            if cursor.fetchone():
                temp_name = f"__temp_{time.time()}__"
                cursor.execute("INSERT INTO contacts (avatar_hash, name_img_hash, nickname) VALUES (?, ?, ?)", (avatar_hash, name_img_hash, temp_name))
                new_id = cursor.lastrowid
                final_nickname = f"{raw_nickname}_{new_id}"
                cursor.execute("UPDATE contacts SET nickname = ? WHERE id = ?", (final_nickname, new_id))
            else:
                final_nickname = raw_nickname
                cursor.execute("INSERT INTO contacts (avatar_hash, name_img_hash, nickname) VALUES (?, ?, ?)", (avatar_hash, name_img_hash, final_nickname))
                new_id = cursor.lastrowid
            conn.commit()
            return new_id, final_nickname
        finally:
            conn.close()
    
    def save_message(self, contact_id, contact_name, sender, content):
        conn = self._get_connection()
        try:
            timestamp = int(time.time())
            cursor = conn.cursor()
            cursor.execute("INSERT INTO messages (contact_id, contact_name, sender, content, timestamp) VALUES (?, ?, ?, ?, ?)", 
                           (contact_id, contact_name, sender, content, timestamp))
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()
    
    def get_context(self, contact_id, limit=10):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT sender, content FROM messages WHERE contact_id = ? ORDER BY timestamp DESC LIMIT ?", (contact_id, limit))
            results = cursor.fetchall()
            context = [{"role": "assistant" if s == "我" else "user", "content": c} for s, c in reversed(results)]
            return context
        finally:
            conn.close()

    # --- 以下为给 GUI 专属定制的数据读取方法，杜绝 GUI 直触 SQL ---
    def get_all_messages_for_export(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT contact_name, sender, content, timestamp FROM messages ORDER BY timestamp DESC")
            return cursor.fetchall()
        finally:
            conn.close()
            
    def get_chat_history_by_name(self, contact_name):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM contacts WHERE nickname = ?", (contact_name,))
            res = cursor.fetchone()
            if res:
                cursor.execute("SELECT sender, content, timestamp FROM messages WHERE contact_id = ? ORDER BY timestamp ASC", (res[0],))
            else:
                cursor.execute("SELECT sender, content, timestamp FROM messages WHERE contact_name = ? ORDER BY timestamp ASC", (contact_name,))
            return cursor.fetchall()
        finally:
            conn.close()
            
    def get_latest_messages_all_contacts(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT id, nickname FROM contacts")
            contacts = cursor.fetchall()
            results = []
            for cid, cname in contacts:
                cursor.execute("SELECT content FROM messages WHERE contact_id = ? AND sender = 'user' ORDER BY timestamp DESC LIMIT 1", (cid,))
                u_msg = cursor.fetchone()
                cursor.execute("SELECT content FROM messages WHERE contact_id = ? AND sender = '我' ORDER BY timestamp DESC LIMIT 1", (cid,))
                b_msg = cursor.fetchone()
                if u_msg or b_msg:
                    results.append({
                        "contact_name": cname,
                        "user_message": u_msg[0] if u_msg else "",
                        "bot_message": b_msg[0] if b_msg else ""
                    })
            return results
        finally:
            conn.close()

    def clear_all_messages(self):
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM messages")
            conn.commit()
        finally:
            conn.close()
            
    def close(self): pass
