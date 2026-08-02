# Office-room-set-up-



This project was done for the office for student involvement at Fordham University, where I work as a student worker. Our office is responsible for all the student-run and external events across campus. We, as student workers, are tasked with unlocking the event spaces, and we also set up any audio/visual equipment, like microphones, projectors, and Zoom calls if needed. To see the events and equipment needed, we use 25live, and then we put all the needed information onto a whiteboard. However, my manager does not want the whiteboard anymore. It takes a long time to set up, and workers often make mistakes. This project aims to make it quicker and easier for the student workers to get access to the information they need. Now we are replacing the whiteboard with a TV, likely connected to a computer. This TV will have all the information the workers need to get their tasks done.

This project will be done in a couple of steps.

First, I want to just test and make sure this works on the TV. The first program I will build is a program that takes a PDF of the day's events and converts it to a excel sheet for the workers to read. This Excel sheet will have the event location, the time the room needs to be set up, and what equipment is needed for the event. The idea is that the worker can log in to 25live, download the pdf which can be confusing and hard to read, and drop the file into the program where it outputs the important info into an Excel sheet, making it easier to read.

After testing and making sure this bare-bones approach works, I want to fully automate this process. How to do that, I'm not fully sure, but 25live does have an api so maybe I can just make a program that calls the api and puts the info into a spreadsheet. Or my second idea was to maybe set up an AI agent to help with this. Fordham has access to Google's Gemini Pro, so might be able to set something up with that. Will have to do more research.
