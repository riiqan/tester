import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'notes.db')

def update_database():
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Check existing columns
        cursor.execute("PRAGMA table_info(notes)")
        columns = [col[1] for col in cursor.fetchall()]
        
        # Add new columns if they don't exist
        new_columns = {
            'reading_time': 'TEXT',
            'difficulty_level': 'TEXT',
            'prerequisites': 'TEXT',
            'learning_objectives': 'TEXT'
        }
        
        for col_name, col_type in new_columns.items():
            if col_name not in columns:
                cursor.execute(f"ALTER TABLE notes ADD COLUMN {col_name} {col_type}")
                print(f"✅ Added column: {col_name}")
            else:
                print(f"ℹ️ Column already exists: {col_name}")
        
        conn.commit()
        conn.close()
        print("\n✅ Database updated successfully!")
        
    except Exception as e:
        print(f"❌ Error updating database: {e}")

if __name__ == "__main__":
    update_database()