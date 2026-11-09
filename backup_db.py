from app import init_db, create_database_backup

if __name__ == "__main__":
    init_db()
    path = create_database_backup(force=True)
    print(f"Backup created: {path}")
