# VBA-challenge
VBA Homework for ToddT

I have completed the Moderate VBA assignment, but I started to create the column headers for the #challenge. 


For the segment below, I was trying to figure out a way to get rid of the "NaNs" but ran into an issue, so the 0's are still in. 

    'Calculate the percent change and write to the appropriate cell and set result to two decimals
        If open_year = 0 Then
            ws.Cells(Summary_Table_Row, 11).Value = "NaN"
        Else
            ws.Cells(Summary_Table_Row, 11).Value = ((ws.Cells(i, 6).Value - open_year) / open_year)
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        End If
        
'Sets the opening year value to the new value
    open_year = ws.Cells(i + 1, 3).Value
    

What would be the best way to clear all percent changes that are zero? I ran out of time on this assignment, but I'm not sure the best place/what code to put to do that action.

