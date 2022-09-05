# VBA Challenge


## Background

The main idea is to prepare a VBA code for client, Steve, in order to assist him in analysing returns from the various stocks in the year 2017 and 2018. This analysis will help Steve guide his parents towards a good investment decision.

### Purpose

The purpose of the project is to edit or refactor the existing VBA code used for the analysis of various stocks in the year 2017 and 2018. The main reason behind this is to make the code more efficient by reducing the running time.

## Results

### Refactoring the Code

The procedure for refactoring the code included the creation of three output arrays, 'tickerVolumes', 'tickerStartingPrices' and 'tickerEndingPrices', in addition to the array 'tickers'. The variable 'tickerIndex' will be used as a variable to iterate over all the rows.

### Refactored Code

'''

 
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
         
         Dim tickerIndex As Integer
       ' Initiate tickerIndex at zero
         tickerIndex = 0
    
    
    '1b) Create three output arrays
    
          Dim tickerVolumes(12) As Long
          Dim tickerStartingPrices(12) As Single
          Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
          
          For i = 0 To 11
         
         ' Initiate each tickervolume at zero
            tickerVolumes(tickerIndex) = 0
         
         
          Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
         
           For j = 2 To RowCount
    
     '3a) Increase volume for current ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        'If  Then
            
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
             tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
             ' End If
           End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If  Then
              
              If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
             
                  tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
        
           End If
            
        '3d Increase the tickerIndex.
        
               If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
             
                 tickerIndex = tickerIndex + 1
                 
                 End If
            
              Next j
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
               
               For i = 0 To 11
           
           ' Activate Output Worksheet
               Worksheets("All Stocks Analysis").Activate
         
          ' Ticker Row Label
              Cells(4 + i, 1).Value = tickers(i)
           
           ' Sum of Volume
             Cells(4 + i, 2).Value = tickerVolumes(i)
         
           ' Return Value
             Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
             
            Next i
            
  '''
  
### The original code

'''


       Dim tickers(12) As String

       tickers(0) = "AY"
       tickers(1) = "CSIQ"
       tickers(2) = "DQ"
       tickers(3) = "ENPH"
       tickers(4) = "FSLR"
       tickers(5) = "HASI"
       tickers(6) = "JKS"
       tickers(7) = "RUN"
       tickers(8) = "SEDG"
       tickers(9) = "SPWR"
       tickers(10) = "TERP"
       tickers(11) = "VSLR"
    
     '3a) Initialize variables for starting price and ending price
       
          Dim startingPrice As Double
          Dim endingPrice As Double
        
     '3b) Activate data worksheet
          
          Worksheets(yearValue).Activate
          
     '3c) Get the number of rows to loop over
         
         RowCount = Cells(Rows.Count, "A").End(xlUp).Row

         
      '4) Loop through tickers
     
        For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
       
     '5) loop through rows in the data
   
        Worksheets(yearValue).Activate
       
         For j = 2 To RowCount
            
            '5a) Get total volume for current ticker
           
             If Cells(j, 1).Value = ticker Then

            totalVolume = totalVolume + Cells(j, 8).Value

            End If
           
        '5b) get starting price for current ticker
           
             If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

             startingPrice = Cells(j, 6).Value

             End If
 
         '5c) get ending price for current ticker
           
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           endingPrice = Cells(j, 6).Value
           
           End If


        Next j
       
    '6) Output data for current ticker

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   
         Next i
         
 '''

The idea of using arrays will help to reduce the running time of the code and make it more efficient as compared to the original one.

## Elapsed run time of Original Code

### Run time for the year 2017
<img width="1440" alt="2018 original" src="https://user-images.githubusercontent.com/111387025/188322942-bbf17e6d-50fe-4858-9420-7aa4c7ecdd61.png">

### Run time for the year 2018

<img width="1440" alt="2018 original" src="https://user-images.githubusercontent.com/111387025/188323893-ca9bb9b5-0ecf-469c-8787-c92c8d04e167.png">

##Elapsed run time of Refactored Code

### Run time for the year 2017
<img width="1440" alt="2017 Refactored" src="https://user-images.githubusercontent.com/111387025/188323072-33b956be-adf5-4f43-b041-9f097fe9d561.png">

### Run time for the year 2018
<img width="1440" alt="2018 Refactored" src="https://user-images.githubusercontent.com/111387025/188323085-b20ca03b-9698-4e79-91b2-326b778f9b10.png">

## Original Code Vs. Refactored Code

By refactoring the code it can be clearly seen that it is much faster than the original one. The original code took approximately 0.4 seconds to run however, the refactored code just took 0.08 seconds to run, making it 5 times quicker than the former for the both years.

## Summary 

### Advantages and Disadvantages of Refactoring

The advantage of refactring the code is that it reduces the running time of the code and helps in making the code efficient. On the other hand, major drawback is editing the code which is already working well. This sometimes can lead to code which migh not run at all or the wasting the original code which was working.

### Advantages and Disadvantages: Original and Refactored VBA Script

VBA gives privilege of using the old and the new code simultaneously, which makes the process of refactoring easier. However, the downside towards it can be the language used i.e. Syntax, as the code can be perfectly correct but the use of incorrect language might not let VBA run the code ending in some error.

 
