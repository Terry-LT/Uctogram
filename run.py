import sys
from bot import run

if __name__ == "__main__":
    backend = sys.argv[1] if len(sys.argv) > 1 else "local"
    run(backend)