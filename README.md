# Betting Excel List

This application retrieves betting matches using the Private Betting API and saves them into an Excel file. Users can specify date options to fetch matches for today, tomorrow, or the day after tomorrow and convert them into an Excel file.

## How to Use

1. Clone or download this project.
2. Make sure Python 3.x is installed.
3. Install the required libraries by running the following command in your terminal or command prompt:
    ```
    pip install requests xlsxwriter
    ```
4. Run the `BetPopulerMatches.py` file in your terminal or command prompt.
5. Choose an option according to the instructions provided:
    - 0: All
    - 1: Today
    - 2: Tomorrow
    - 3: Day after tomorrow

## Example

```bash
python BetPopulerMatches.py
