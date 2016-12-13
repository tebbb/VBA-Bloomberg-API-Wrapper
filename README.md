# VBA-Bloomberg-API-Wrapper

A VBA Class that enables you to querry a wide range of Bloomberg Data with simplicity. Handles errors that Bloomberg might throw at you. 
You can request Reference Data, Historical Data and for those with an AIM subscription: Positions Data.

# REQUIREMENTS
Bloomberg Desktop installed with Excel Add-in installed
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

# GETTING STARTED
In VBA (Alt+F11):
  Insert a Module
  Insert the following Code:

'''
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
    
    Dim i as integer, j as integer, PrintStr as string
    for i=1 to ubound(BData,1)
      PrintStr=""
      for j=0 to ubound(BData,2)
        PrintStr=PrintStr & " " & BData(i,j) 
      next j
      debug.print PrintStr
    next i
  
End Sub'''

# CODE STRUCTURE


Code for Excel VBA and Bloomberg Desktop API
Need to unable in Tools->References: Bloomberg API COM 3.5 Type Library
Insert the Code in a new Class and you are ready to go!

