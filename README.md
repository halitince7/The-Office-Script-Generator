# The Office Script Generator

A Python application that generates episode scripts from "The Office" TV show using a comprehensive dialogue database. Available in both GUI and command-line versions.

## Features

- Two interfaces:
  - GUI application with easy-to-use controls
  - Command-line tool for quick access
- Season and episode selection
- Formatted Word document output with:
  - Episode headers
  - Character names in bold
  - Clean line spacing
  - Professional margins
- Automatic saving to current directory (CLI) or Downloads folder (GUI)
- Clear status updates and error handling


## Installation

1. Clone the repository
2. Install requirements:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Version

1. Run the GUI application:
```bash
python office_gui.py
```

2. Select a season from the dropdown menu
3. Enter episode numbers separated by commas (e.g., "13,14,15")
4. Click "Generate Scripts"
5. The formatted document will be saved to your Downloads folder

### Command Line Version

Run the script with season and episode numbers as arguments:
```bash
python office.py <season> <episode_numbers>
```

Examples:
```bash
# Generate script for Season 5, Episode 13
python office.py 5 13

# Generate script for Season 5, Episodes 13, 14, and 15
python office.py 5 13,14,15
```

Output files will be saved to your Downloads folder with names like:
- `office_s05e13.docx` (for season 5, episode 13)
- `office_s05e13-14-15.docx` (for season 5, episodes 13, 14, and 15)

## File Naming

- Both GUI and CLI versions save to Downloads folder: `~/Downloads/office_s05e13.docx`

## Data Source

The application uses `The-Office-Lines-V4.csv`, which contains all dialogue from The Office (US) TV series, sourced from [Kaggle: The Office (US) Complete Dialogue/Transcript](https://www.kaggle.com/datasets/nasirkhalid24/the-office-us-complete-dialoguetranscript). The dataset includes:
- Season and episode numbers
- Scene numbers
- Speaker names
- Dialogue lines
- Episode titles

## File Structure

```
.
├── office_gui.py          # GUI application
├── office.py             # Command-line tool
├── The-Office-Lines-V4.csv # Dialogue database
└── README.md
```

## Notes

- The CSV file must be in the same directory as the script
- Both GUI and CLI versions save output files to Downloads folder
- Both versions provide status updates and error messages

## Error Handling

Common errors and solutions:
- "File not found": Ensure The-Office-Lines-V4.csv is in the same directory
- "Invalid season/episode": Check that your input matches available episodes
- "Module not found": Run pip install for required packages

## License

This project is for educational purposes only. The Office is owned by NBC Universal.
```