from pathlib import Path

# Cartella del progetto (dove sta questo file)
PROJECT_DIR = Path(__file__).parent.absolute()

# Database
DB_DIR = PROJECT_DIR / "data"
DB_DIR.mkdir(exist_ok=True)
DB_PATH = DB_DIR / "materie_prime.db"

# Esportazioni
EXPORT_DIR = PROJECT_DIR / "exports"
EXPORT_DIR.mkdir(exist_ok=True)
