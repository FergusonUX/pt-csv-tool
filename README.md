# pt-csv-tool
Converting SM9 Feedback export files into something better.

This document serves as a guide to converting the IT FOL report from it’s current (2.22.18) format into the new format.  

## Requirements
**Windows**: Command Prompt
**Mac**: Terminal

## Installation
### 1. Install Node
This converter utility is written in Node.js. You must download and install Node on your computer first to run the program.
[Download Node here](https://nodejs.org/en/download/).
### 2. Download program
We keep the most up-to-date version of the program on Google Drive. You can download it here. Alternatively, you can view the folder contents. Download it and unzip it wherever you like (ex: Desktop).
### 3. Install Dependencies
Launch command-line tools
**On Windows**: Start Menu > Search > Command Prompt
**On Mac**: Cmd+Spacebar > Terminal

At the prompt, enter: `cd PATH/TO/FOLDER`
If you don’t know the path, you can drag the folder on to the command prompt and it will place the path for you.
Next, enter: `npm install`
This will install the dependencies for the program to run.

## Running the Program
### 1. Save your spreadsheet to the program folder
### 2. Execute
At the command prompt, enter: `node converter -i YOURFILENAME`
**Note**: Do not include the extension in the filename.
