# Research-Services

Project 1:

   Analyzing Metrics for the department of Research Services at UPEI.
   
   Language: Google Apps Script (Javascript)
   
   OOP - Emulating an OOP project architecture in a functional language
   
   1. Think of the trigger as a "main" function. This is what is executed every time an edit happens.
   2. Use "initialize" as a constructor for building the current "client" and "quarter" objects
      - Client is like a regular object, containing data
      - Quarter is like a parent object containing all the individual quarters, which themselves contain the relevant data for their quarter
   3. The other utility functions are just for simplifying the code

Project 2:

   Google Apps Script for batch updating Google Drive folder names.

   Broken into 2 functions - 

   1. main - defines drive folders by "ID"
   2. rename - Executed in main, implementing "setName()" on an input folder ID
