[Chinese Version](README.md)

# Motor Lottery

This project is a Python application for conducting a motor parking space lottery for 文華天際 community. It uses a graphical interface built with `tkinter` and integrates `xlwings` to update an Excel file in real-time, so the lottery results are immediately reflected in an open Excel sheet.

## Features
- Allows users to conduct lotteries for different buildings (e.g., A棟, B棟) in the community.
- Automatically draws participants and assigns parking spaces.
- Results are written directly to an Excel file in real-time, allowing for instant updates if the file is open.

## Demo Video
[Demo Video Link](https://drive.google.com/file/d/1J9sSs46bFs52iwFfptUgivsCy5NFHmKE/view?usp=sharing)

## Requirements
- Python 3.x
- Libraries:
  - `openpyxl` (for Excel file manipulation)
  - `xlwings` (for real-time Excel updates)
  - `tkinter` (for GUI, comes pre-installed with Python)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/leo851001/motor_lottery.git
   cd motor_lottery
   ```

2. Install the required libraries:
   ```bash
   pip install openpyxl xlwings
   ```

## Usage

1. Run the application:
   ```bash
   python motor_lottery_gui.py
   ```
2. Select either "A棟抽籤" or "B棟抽籤" to start the lottery for the chosen building.
3. After the lottery for a building is completed, you can immediately start the lottery for the other building by clicking the corresponding button.

## Notes

- Ensure that the Excel file remains open to view real-time updates.
- The application will automatically update the results in the designated Excel file.

