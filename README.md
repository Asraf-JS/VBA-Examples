# VBA-Examples
This Github repository is a collection of VBA (Visual Basic for Applications) code examples and snippets for Excel and other Microsoft Office applications. The examples cover a range of use cases and functions, from automating repetitive tasks to data analysis and reporting.

**extractNumbers**

In this code, we're using a For loop to iterate through each character in the input string (inputString). We use the IsNumeric function to check if each character is a number. If the character is a number, we append it to the outputString variable. Finally, we write the outputString to cell B1.

Note that this code assumes that the input string contains only alphanumeric characters and numeric digits. If the input string contains other types of characters, such as symbols or special characters, those characters will be excluded from the output string.

**extractCityAndState**

Note that this code assumes that the state keyword is present in the address and that the keyword is spelled correctly. If the address does not contain a state keyword or the keyword is spelled incorrectly, the city and state variables will be left blank. Additionally, there may be other approaches to identifying the state in an address, depending on the specific structure and content of the address data.

**changeTextCase**

If the input text is in all caps, we use the LCase function to convert it to all lowercase. If the input text is in title case or mixed case, we use the StrConv function with the vbProperCase argument to convert it to title case. If the input text is already in all uppercase, we leave it unchanged.

**SimpleAI**

In this example, the SimpleAI macro implements a simple decision tree algorithm that predicts whether or not someone should play golf based on the humidity level and wind conditions. The decision tree is represented as a two-dimensional array, where each row represents a node in the tree and the columns represent the decision features and outcomes.

**Combine_Sheets**

The VBA code provided combines data from two worksheets, Sheet1 and Sheet2, into a new worksheet called "CombinedSheet." You can find the files used in Sample Files folder called "Combined Sheets"
