# CSV Tool
Converting SM9 Feedback export files into something better.

This document serves as a guide to converting the IT FOL report from it’s current format into the new format.  

## Requirements
**Windows**: Command Prompt  
**Mac**: Terminal

## Installation
### 1. Install Node
This converter utility is written in Node.js.  
You must download and install Node on your computer first to run the program.
[Download Node here](https://nodejs.org/en/download/).
### 2. Download Repo
Download [the repo](https://github.com/FergusonUX/pt-csv-tool/archive/master.zip).
### 3. Install Dependencies
Open your command-line tool of choice:  
**On Windows**: Start Menu > Search > Command Prompt  
**On Mac**: Cmd+Spacebar > Terminal

At the prompt, enter: `cd PATH/TO/REPO/FOLDER`  
If you don’t know the path, you can drag the folder on to the command prompt and it will place the path for you.  
Next, enter: `npm install`  
This will install the dependencies for the program to run.  

## Running the Program
### 1. Place a spreadsheet to be converted in the script folder
### 2. Run script
At the command prompt, enter: `node converter -i YOURFILENAME`
**Note**: Do not include the extension in the filename.
