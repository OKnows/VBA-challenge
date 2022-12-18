# VBA Homework: The VBA of Wall Street

First I declared the variables I would be using throughout the exercise.

Then I created new column headers for the ticker, yearly change, percentage change, and total volume.
    And placed a spot for the Greatest % increase and decrease and Greatest total volume
    
I created the ticker counter as we go through rows 2 to last row.
 Then using an if statement, we specify that if the ticker symbol changes in the next row, then we would add the symbol, total volume, yearly change, and percentage change to the specified locations.
 
At this point it conditionally formats the negative and positive changes so that if the change is greater than 0, it is green, if it's less than 0 it's red, and if it's 0 it isn't formatted.

Then the code has it going to the next ticker and resetting the value to 0.

However if the ticker symbol in the next row is the same as the one previously, we add the volume in the volume column to a total variable each time it's the same.

For the bonus, we are going through the entire Percentage change column and finding the maximum and minimum values and adding that value and its corresponding ticker to the right of its label.
Then it does the same thing for greatest volume by going down the range of the total volume column and placing that with the corresponding ticker next to the label.
We went back through and added ws to the script to read through all of the worksheets.



