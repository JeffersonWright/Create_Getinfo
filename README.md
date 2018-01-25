CreateTextjoin.gs is a Google Script that takes a blank google sheet and turns it into an interface used to join and interact with strings in multiple cells.
Once created, this sheet will take all the entries in columns A through D and join them into a single string, with the output being cell P3.
I created this from scratch while working at TestBest doing data collection and it greatly increased the team's productivity.
This function needs to be run once and then the program can be used.
To create this function, open a new Google Sheet, click Tools>Script Editor, and paste the entire code in CreateTextjoin.gs over the default code. Then hit "Run" or the run arrow icon and the new sheet will be transformed into the textjoin sheet.


--Instructions For Use--

Input:
Anything entered into cells in colums A through D is the input, with blank cells being ignored.
The cells will be joined with whatever characters are in cell F3.
Input rows can be skipped by entering into cell F4, for example "Skip Lines: 2" will capture every third line in columns A through D.
Cell N3 can be changed to select only cells that begin with a certain string, for example "Starts with: Gr" would exclude all cells that do not start with "Gr".

Modification:
The output string will be displayed in the large box starting in cell E5.
The output string can be modified by entering into the blank cells in F1:L2. There are four chances to remove a string and replace it with a different string, labelled accordingly.

Output:
Cell H3 displays the total characters of the input text.
Cell K3 displays how many characters were removed by the four Remove Char strings.
Cell K4 displays how many characters were replaced by the Replace With strings.
Cell H4 displays the total characters of the output text.
Once you are satisfied with your output, copy cell P3. When pasting into your destination, be sure to paste without formatting (Ctrl+Shift+V). If you forget this, what will be pasted is the formula and a yellow background to remind you to paste without formatting.
