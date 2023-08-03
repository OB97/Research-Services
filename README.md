# Research-Services
Working on Google Sheets/Excel files to gather and analyze Metrics for the department of Research Services at UPEI.

Language: Google Apps Script (Javascript)

OOP - Emulating an OOP project architecture in a functional language

1. Think of the trigger as a "main" function. This is what is executed every time an edit happens.
2. Use "initialize" as a constructor for building the current "client" and "quarter" objects
   i. Client is like a regular object, containing data
   ii. Quarter is like a parent object containing all the individual quarters, which themselves contain the relevant data for their quarter
3. The other utility functions are just for simplifying the code
