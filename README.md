# VBA Bloomberg API Wrapper

A VBA Class that enables you to querry a wide range of Bloomberg Data with simplicity. Handles errors that Bloomberg might throw at you. 
You can request Reference Data, Historical Data and for those with an AIM subscription: Positions Data.

# REQUIREMENTS
    Bloomberg Desktop installed with Excel Add-in installed (and a valid subscription)
    In VBA (Alt+F11) enable in Tools->References: Bloomberg API COM 3.5 Type Library
    In order to get the data you will need to be logged in you Bloomberg account

# HOW TO INSTALL
    In VBA (Alt+F11):
    Create a new Class and call it (for example): C_Bloom
    Insert the code
    Done Installing!

# WHERE TO FIND FIELDS
    The Fields relating tot he data you are looking for can be found form your terminal.
    1. Go to a ticker you are interested in
    2. Type the function FLDS in your terminal
    3. Search the information you are looking for

# GETTING STARTED EXAMPLES
    In VBA (Alt+F11):
    Insert a Module
    Insert the following Code:

```
'--> Reference Data <--
Public Sub TestBloom()
    Dim Bloom as new C_Bloom  'Assuming you called your wrapper class C_Bloom
    Dim BData as Variant      'What will collect the data querried
    Dim Tickers() as String   'A List of Bloomberg Tickers starting at rank 1
    Dim Fields() as String    'A List of Fields starting at rank 1 

    Redim Tickers(1 to 3),Fields(1 to 2)
    Tickers(1)="USDEUR Curncy"
    Tickers(2)="AAPL US Equity"
    Tickers(3)="GT10 Govt"
    Fields(1)="PX_LAST"
    Fields(2)="CRNCY"
    
    BData=Bloom.referenceData(Tickers,Fields) 'Your data is now in BData
    
    'Print the Data in the Immediate Window to see what we have (or explore BData yourself in the Locals Window)
    Dim i as integer, j as integer, PrintStr as string
    for i=1 to ubound(BData,1)
        PrintStr=""
        for j=0 to ubound(BData,2)
            PrintStr=PrintStr & " " & BData(i,j) 
        next j
        Debug.Print PrintStr
    next i
    
    'Output should similar to:
    'USDEUR Curncy 0.9426 EUR
    'AAPL US Equity 113.3 USD
    'GT10 Govt 95.96875 USD
End Sub
```
```
'--> Historical Data <--
Public Sub TestBloom()
    Dim Bloom As New C_Bloom
    Dim BData As Variant
    Dim Tickers() As String
    Dim Fields() As String
    Dim StartD As Date
    Dim EndD As Date
    
    ReDim Tickers(1 To 3), Fields(1 To 2)
    Tickers(1) = "USDEUR Curncy"
    Tickers(2) = "AAPL US Equity"
    Tickers(3) = "GT10 Govt"
    Fields(1) = "PX_LAST"
    Fields(2) = "CHG_PCT_1D"
    StartD = Date - 20
    EndD = Date
    
    BData = Bloom.historicalData(Tickers, Fields, StartD, EndD) 'Your data is now in BData
    
    Dim i As Integer, j As Integer, PrintStr As String
    For i = 1 To UBound(BData, 1)   'Going through the Tickers
        Debug.Print BData(i, 0)
        For k = 1 To UBound(BData(i, 1), 1) 'Going through the Dates
            PrintStr = ""
            For j = 1 To UBound(Fields) + 1 'Going through the Fields
                PrintStr = PrintStr & "  " & BData(i, j)(k)
            Next j
            Debug.Print PrintStr
        Next k
    Next i
    
End Sub
```
```
'--> Positions Data <--
Public Sub TestBloom()
    Dim Bloom As New C_Bloom
    Dim BData As Variant
    Dim accountType As String   'Account, Group, ...
    Dim accountName As String   'The Name of the account/Fund
    Dim Fields() As String
    
    ReDim Fields(1 To 3)    'Note that there are two extra fields returned in front: Ticker and Name
    Fields(1) = "POS_CN"
    Fields(2) = "CNTRY_OF_RISK"
    Fields(3) = "YELLOW_KEY"
    
    accountType = "Account"
    accountName = "SOME_ACCOUNT_NAME"
    
    BData = Bloom.PortfolioPositionData(accountType, accountName, Fields)
    
    Dim i As Integer, j As Integer, PrintStr As String
    For i = 1 To UBound(BData, 1)   'Going through the positions
        PrintStr = ""
        For j = 1 To UBound(BData, 2)   'Going through the Fields
            PrintStr = PrintStr & "  " & BData(i, j)
        Next j
        Debug.Print PrintStr
    Next i
    
End Sub
```

# CODE STRUCTURE
I. Three accessible Functions for each type of data request:
        referenceData,
        historicalData,
        PortfolioPositionData. 
    They take the inputs and check for errors, convert dates to the Bloomberg format, call the general sub for data request and return   the data once it is ready.

II.  The ProcessDataRequest and its 3 dependent Functions:
        ProcessDataRequest,
        OpenService,
        SendRequest,
        catchServerEvent. 
    ProcessDataRequest is called by the accessible functions. It coordinates the different steps of a request: opening a service (with OpenService), sending a request (with SendRequest), and listening for an answer from Bloomberg (with catchServerEvent). If any error occurs the function will return false to ProcessDataRequest and it can close the object cleanly and raise an appropriate error.
    
III. Three server data processing functions:
        getServerData_reference,
        getServerData_historical,
        getServerData_portfolio. 
    These will be called by the catchServerEvent depending on the request sent. They purpouse is to structure the data for output and catch errors returned by Bloomberg.
    
IV. The Data returned: is in a matrix of type variant in order to support the different data types returned by Bloomberg

