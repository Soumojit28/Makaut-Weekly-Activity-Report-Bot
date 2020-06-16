# Makaut-Weekly-Activity-Report-Bot
### A bot written in python to automate Weekly Activity Report submission form a excel spread sheet file.

![](https://i.ibb.co/hFrztnV/header.png)

## Installtion
- First download the latest Bot relese from [here]()
- Download the chromedriver from [here](https://chromedriver.chromium.org/). Make sure to donwload the appropiate version to match with the browser version. To check the browser version goto chrome->menu->help->about
- Put the bot, chrome driver and excel spreadsheet in a folder and start the bot.
- follow on screen instructions.

## Spreadsheet Specification
The spreadsheet muct be in this specified format-
```
Coulumn -> Data
A -> Blank
B -> Week
C -> Date(Optional Bot will not read from this)
D -> Topic
E -> Platform
F -> Time(Optional Bot will not read from this)
G -> Link
H -> Duration
I -> Note
J -> Assignment Received
K -> Assignment Submitted
L -> Test

```
![](https://i.ibb.co/fCLk9jm/Screenshot-201.png)

## Usage
- Put all three files (bot, web driver and spreadsheet) in a same folder
- Then start the bot, provide your login information and continue on screen instruction
- One spread sheet must contain only a specific subject info

## Release History

* 0.1.0
   * Made single exe file
   * Added support for auto user login
   * Added support for all semester and departments
* 0.0.5
   * Created the first test version
    



