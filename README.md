# Flight Booking Bot

A Telegram bot for booking flights, with automatic Excel file integration for storing flight information and user bookings.

## Features

### User Features
- View available flights
- Book flights with an interactive menu system
- View personal bookings
- Cancel bookings

### Admin Features
- Add new flights
- Remove flights
- Download booking data
- Regenerate sample Excel data

## Installation

1. Clone this repository
   ```bash
   git clone https://github.com/yourusername/flight-booking-bot.git
   cd flight-booking-bot
   ```

2. Install required dependencies
   ```bash
   pip install python-telegram-bot pandas openpyxl
   ```

3. Configure the bot
   - Replace `YOUR_TOKEN_HERE` with your Telegram bot token in `main.py`
   - Update the `ADMIN_USERS` list with Telegram user IDs of administrators

4. Run the bot
   ```bash
   python main.py
   ```

## Usage

### Regular User Commands

| Command | Description |
|---------|-------------|
| `/start` | Start the bot and display available commands |
| `/book` | Begin the flight booking process |
| `/my_bookings` | View your current bookings |
| `/cancel_booking` | Cancel one of your bookings |
| `/available` | Show all available flights |

### Admin Commands

| Command | Description |
|---------|-------------|
| `/add_flight` | Add a new flight |
| `/remove_flight` | Remove an existing flight |
| `/download_bookings` | Download the bookings Excel file |
| `/recreate_excel` | Regenerate Excel files with sample data |

## Excel File Structure

The bot automatically creates and maintains two Excel files:

### flights.xlsx
Contains all flight information:
- Date
- Time
- Flight Number
- Departure
- Destination
- Capacity
- Booked seats

### bookings.xlsx
Contains all booking information:
- Date
- Time
- Flight Number
- User ID
- Username
- Booking Time

## User Guide

### How to Book a Flight
1. Send `/book` to start the booking process
2. Select an available date from the menu
3. Choose a flight from the available options
4. Confirm your booking
5. You'll receive a confirmation message when successful

### How to Cancel a Booking
1. Send `/cancel_booking` to start the cancellation process
2. Select the booking you wish to cancel from the menu
3. The bot will confirm when your booking has been canceled

### How to View Your Bookings
Send `/my_bookings` to see a list of all your active bookings

## Admin Guide

### Adding a New Flight
1. Send `/add_flight` to start the process
2. Enter flight details in the format:
   ```
   YYYY-MM-DD HH:MM FlightNumber Departure Destination Capacity
   ```
   Example: `2025-04-01 08:30 FL123 New_York London 120`

### Removing a Flight
Send `/remove_flight YYYY-MM-DD HH:MM FlightNumber`

Example: `/remove_flight 2025-04-01 08:30 FL123`

Note: Flights with existing bookings cannot be removed until all bookings are canceled

### Downloading Booking Data
Send `/download_bookings` to receive the complete bookings Excel file

### Regenerating Sample Data
Send `/recreate_excel` to generate fresh sample flight data

Warning: This will erase all existing flight and booking information!

## Technical Details

- Built with python-telegram-bot v20+
- Uses pandas and openpyxl for Excel manipulation
- Implements conversation handlers for interactive booking flows
- Automatically creates necessary Excel files on first run
- Includes sample data generator for testing

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgements

- [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot) library
- [pandas](https://pandas.pydata.org/) for data manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel integration
