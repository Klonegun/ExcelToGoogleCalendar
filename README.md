
# Excel to Google Calendar

Ah, the eternal struggle against the mighty force of laziness! Picture this: I'm at my desk, staring down the daunting Excel spreadsheet that holds my work schedule, while my dear wife kindly requests its migration to our shared Google Calendar. Cue the eye-roll and the sighs of resignation. But fear not! Rather than succumbing to the tedious chore, I decided to let technology do the heavy lifting. And thus, this ingenious program was born to liberate me from the clutches of manual labor!




## API Reference

#### Google Calendar API

```http
  https://developers.google.com/calendar/api/v3/reference
```

| Parameter | Type     | Description                |
| :-------- | :------- | :------------------------- |
 `SCOPES` | `string` | **Required**. Authorization Scope added to the Google Cloud project. |
  `Credentials / Creds` | `.JSON` | **Required**. Allows access to API through project. |
   `CalendarID` | `string` | **Required**. This ID will point to your Calendar. |


## Coding Language

Python 3
## Libaries

External Libaries used:

```bash
  ttkbootstrap - Used for GUI
  openpyxl - Used to read Excel Spreadseet
  google
  googleapiclient
```
    
## Features

- Application reads Excel file and feeds values into the functions.
- Ability to find different users based on user input.
- Breaks down strings imported from Excel into useable data for Google.
- Automatically uploads calendar events to Google Account. 


## Screenshots

![App Screenshot]([file:///C:/Users/Ky%20Farrar/PycharmProjects/ConverterScreenShots.pdf](https://github.com/Klonegun/ExcelToGoogleCalendar/blob/main/ConverterScreenShots.pdf))


## Feedback

If you have any feedback or would like to discuss the project, please reach out to me at kyfarrar@outlook.com

