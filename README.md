# VBA-Challenge
For this challenge we were asked to create a loop in VBA that would return the ticker symbol, the yearly change, the percent change, and the total stock volume of the stock. 
To start with in the VBA code I listed out all my different variables and the value assigned to the variables when necessary.
I listed out every variable I would need in order to return the required information for the columns we had to create: ticker symbol, the yearly change, the percent change, and the total stock volume. 

I used two IF statements in my code, the first IF statment starts with the equation: If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then, and the second IF statement starts with the equation: If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

This statement makes it so that when I run my code excel knows to look at when the cells in the ticker column (Column A) changes ticker symbols. This tells excel when to change to the next ticker symbol and will return all the information for that specific ticker symbol. 

The first IF statement I used to find my opening price. This makes it so that I can find the first opening price and then the last opening price and then the first opening price for the next ticker symbol. 

Then I went into my second IF statement which is my main IF statement. This IF statement contains the equations for my ticker name, ticker volume, closing price, yearly change, and percent change. The equations are listed below: 

    Ticker_Name = ws.Cells(i, 1).Value
    
    Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
    
    Closing_Price = ws.Cells(i, 6).Value
    
    Yearly_Change = Closing_Price - Opening_Price
    
    Percent_Change = Yearly_Change / Opening_Price
    
Then within this IF statement, I created three more IF statements in order to return the Greatest Increase and Ticker symbol, the Greatest Decrease and Ticker symbol, and the Greatest Total Volume and Ticker symbol. The equations for these IF statements are below.

        If Percent_Change > Greatest_Increase Then
        
            Greatest_Increase = Percent_Change
            
            Greatest_Ticker_Name = Ticker_Name
            
        End If
        
        If Percent_Change < Greatest_Decrease Then
        
            Greatest_Decrease = Percent_Change
            
            Greatest_Decrease_Ticker_Name = Ticker_Name
            
        End If
        
        If Ticker_Volume > Greatest_Ticker_Volume Then
        
            Greatest_Ticker_Volume = Ticker_Volume
            
            Greatest_Ticker_Volume_Name = Ticker_Name
            
        End If

Then I proceeded to finish up the IF statement by assigning the ticker name, ticker volume, yearly change, and percent change to their respective columns. I also number formatted the column for the percent change so it would return a percentage. That piece of the code is listed below:

    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
After this I finished with one more equation for an else within the IF statement and then ended the IF statement by putting END IF and then Next i.

Then there was just one more piece in order to complete all the requirements, I needed to assign the cells for the Greatest Increase and Ticker symbol, the Greatest Decrease and Ticker symbol, and Greatest Ticker Volume and Ticker Name. I number formatted the cells for the Greatest Increase and Greatest Decrease for a percent so that it would return a percentage. The equations for these are below:

    ws.Cells(2, 16).Value = Greatest_Ticker_Name
    ws.Cells(2, 17).Value = Greatest_Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = Greatest_Decrease_Ticker_Name
    ws.Cells(3, 17).Value = Greatest_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = Greatest_Ticker_Volume_Name
    ws.Cells(4, 17).Value = Greatest_Ticker_Volume
    
Then to finish it all off I put next ws so that it would run on the next worksheet and then I ended the code and ran it. 
    
The results for the Greatest Increase, Greatest Decrease, and Greatest Ticker Volume for all three sheets are below. 

2018:    
    	Ticker	Value
Greatest % Increase	THB	141.42%
Greatest % Decrease	RKS	-90.02%
Greatest Total Volume	QKN	1.68954E+12
![image](https://user-images.githubusercontent.com/125215083/225708404-1e7424b1-2da8-4cd8-88eb-f1e1396b85ed.png)

2019: 
	Ticker	Value
Greatest % Increase	RYU	190.03%
Greatest % Decrease	RKS	-91.60%
Greatest Total Volume	ZQD	4.37301E+12
![image](https://user-images.githubusercontent.com/125215083/225708765-42125a76-0008-47ef-af90-da61aa914416.png)

2020:
	Ticker	Value
Greatest % Increase	YDI	188.76%
Greatest % Decrease	VNG	-89.05%
Greatest Total Volume	QKN	3.45296E+12
![image](https://user-images.githubusercontent.com/125215083/225708928-bfbda6f3-a7ff-4e43-a47d-594294178036.png)

To finish everything off, I used conditional formatting on the columns holding the yearly change and the percent change so that the negatives were highlighted in red and the positives were highlighted in green. 
