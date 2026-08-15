import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'notes.db')

def add_collection_column():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Check if column exists
        cursor.execute("PRAGMA table_info(notes)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'collection' not in columns:
            cursor.execute("ALTER TABLE notes ADD COLUMN collection TEXT")
            conn.commit()
            print("✅ Added 'collection' column to notes table")
        else:
            print("ℹ️ 'collection' column already exists")
        
        conn.close()
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    add_collection_column()