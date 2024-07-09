# Golf Course Reservation System

## Project Overview

This project is a Golf Course Reservation System designed to streamline the booking process for golf courses and lessons. The system collects user inputs, verifies and processes the data, and provides a detailed cost breakdown for reservations.

## Features

- User-friendly input prompts for mandatory booking information.
- Options to reserve golf courses or register for lessons.
- Conditional fee calculation based on user status (Regular, Senior, Junior) and HKID status.
- Additional service hiring options.
- Detailed billing and confirmation messages.

## File Description

### `golf_course_reservation.xlsm`

This Excel file contains the fee tables and VBA code to manage the golf course reservation system.

### VBA Code Description

#### Global Variables

```vba
Dim Name As String, status As String, strHKID As String
Dim HKID As Boolean
Dim reservation As Date
Dim reserveType As String, lessonType As String, serviceType As String
Dim reservePrice As Integer, lessonPrice As Integer, servicePrice As Integer
Dim code As Integer, hire As Integer
Dim servicePrices As New Collection, serviceTypes As New Collection
Dim tServicePrice As Integer, ferryPrice As Integer
Dim reset As Boolean
```

### Functions and Subroutines
- `resetVariables`: Resets all global variables.
- `userInput`: Collects user information and validates inputs.
- `reserveORregister`: Prompts user to choose between reserving a course or registering for a lesson.
- `reserveCourse`: Manages course reservation process.
- `registerLesson`: Manages lesson registration process.
- `AddService`: Adds selected service details to collections.
- `hireService`: Manages additional service hiring process.
- `finalOutput`: Calculates and displays the total bill.

### Getting Started
To use this system, follow these steps:

1. Clone the Repository:
```bash
git clone https://github.com/your-username/golf-course-reservation.git
```
2. Open the Excel File:
- Navigate to the `golf_course_reservation.xlsm` file.
- Open it with Microsoft Excel.
3. Enable Macros:
- Ensure macros are enabled in Excel to allow the VBA code to run.
4. Run the VBA Code:
- Click on the `Start Booking` button.
