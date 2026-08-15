"""
Database Clearing Script
Run this to delete ALL data from your USP Notes Marketplace database.
WARNING: This is permanent and cannot be undone!
"""

import sqlite3
import os
import sys

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'notes.db')

def clear_database():
    """Delete all data from the database"""
    
    # Check if database exists
    if not os.path.exists(DB_PATH):
        print("❌ Database file not found at:", DB_PATH)
        print("   Nothing to clear.")
        return False
    
    try:
        # Connect to database
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Get list of all tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        
        if not tables:
            print("ℹ️ No tables found in database.")
            conn.close()
            return True
        
        # Clear each table
        for table in tables:
            table_name = table[0]
            if table_name.startswith('sqlite_'):
                continue  # Skip system tables
            
            cursor.execute(f"DELETE FROM {table_name};")
            # Reset auto-increment counters
            cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}';")
            print(f"   ✅ Cleared table: {table_name}")
        
        conn.commit()
        conn.close()
        
        print("\n✅ Database cleared successfully!")
        print(f"   📁 Database: {DB_PATH}")
        return True
        
    except Exception as e:
        print(f"❌ Error clearing database: {e}")
        return False


def delete_database_file():
    """Delete the entire database file (optional, more drastic)"""
    if os.path.exists(DB_PATH):
        try:
            os.remove(DB_PATH)
            print(f"✅ Deleted database file: {DB_PATH}")
            return True
        except Exception as e:
            print(f"❌ Could not delete database file: {e}")
            return False
    else:
        print("ℹ️ Database file doesn't exist.")
        return True


def reset_database():
    """Delete and recreate the database from scratch"""
    print("\n🔴 WARNING: This will delete ALL data permanently!")
    print("   This includes all notes, ratings, and downloads.\n")
    
    response = input("   Are you sure you want to continue? (yes/no): ")
    
    if response.lower() not in ['yes', 'y']:
        print("   ❌ Operation cancelled.")
        return False
    
    confirm = input("   Type 'DELETE ALL' to confirm: ")
    
    if confirm != "DELETE ALL":
        print("   ❌ Operation cancelled. Incorrect confirmation.")
        return False
    
    print("\n🗑️ Clearing database...\n")
    
    # First try to clear all tables
    clear_database()
    
    # Then delete the file entirely
    delete_database_file()
    
    print("\n✅ Database has been completely reset.")
    print("   The database will be recreated automatically when you restart the app.")
    return True


def show_stats():
    """Show current database statistics"""
    if not os.path.exists(DB_PATH):
        print("❌ Database file not found.")
        return
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Get table counts
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        
        print("\n📊 Database Statistics:")
        print("━" * 40)
        
        for table in tables:
            table_name = table[0]
            if table_name.startswith('sqlite_'):
                continue
            
            cursor.execute(f"SELECT COUNT(*) FROM {table_name};")
            count = cursor.fetchone()[0]
            print(f"   📋 {table_name}: {count} records")
        
        conn.close()
        
    except Exception as e:
        print(f"❌ Error getting stats: {e}")


def main():
    """Main menu"""
    print("\n" + "=" * 50)
    print("   USP NOTES MARKETPLACE - DATABASE MANAGER")
    print("=" * 50)
    
    print("\n📁 Database path:", DB_PATH)
    
    if os.path.exists(DB_PATH):
        size = os.path.getsize(DB_PATH)
        print(f"📦 File size: {size/1024:.2f} KB")
    
    print("\nOptions:")
    print("   1. Show database statistics")
    print("   2. Clear all data (keeps tables)")
    print("   3. Complete reset (delete and recreate)")
    print("   4. Exit")
    
    choice = input("\nEnter your choice (1-4): ").strip()
    
    if choice == '1':
        show_stats()
    elif choice == '2':
        clear_database()
    elif choice == '3':
        reset_database()
    elif choice == '4':
        print("👋 Goodbye!")
        sys.exit(0)
    else:
        print("❌ Invalid choice.")
    
    print("\n" + "=" * 50)


if __name__ == "__main__":
    main()