import sys
from office_core import create_script_document

def main():
    # Check command line arguments
    if len(sys.argv) != 3:
        print("Usage: python script.py season episode_numbers")
        print("Example: python script.py 5 13,14,15")
        sys.exit(1)

    season = sys.argv[1]
    episodes = sys.argv[2].split(',')

    # Generate script
    success, message, _ = create_script_document(season, episodes)
    print(message)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main()


