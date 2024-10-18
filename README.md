# VBA Bloomberg API Wrapper

A VBA Class that enables you to querry a wide range of Bloomberg Data with simplicity. Handles errors that Bloomberg might throw at you. 
You can request Reference Data, Port Data, Historical Data. For those with AIM subscriptions I have code to request Positions and Historical Positions data but I am not able to test it at present.

# REQUIREMENTS
    Bloomberg Desktop installed with Excel Add-in installed (and a valid subscription)
    Bloomberg Desktop running while you query data
    

# HOW TO INSTALL
    There are 2 ways to install: either copy the code or install the Add-in
    Copy the Code:
        In VBA (Alt+F11):
            Create a new Class and call it (for example): C_BBG
            Insert the code in C_BBG
            To run the examples: add a Module, and insert the code from Examples
            Ensure that in Tools->References you have Bloomberg API COM 3.5 Type Library selected
    Install the Add-in:
        Download the Add-in
        From your excel File->Options->Add-ins->Excel Add-ins->Go
        Then Browse to where you have downloaded the add-in
        Now VBA (Alt+F11) go to Tools->References and select BBG_Addin in the list

# WHERE TO FIND FIELDS
    The Fields relating tot he data you are looking for can be found form your terminal.
    1. Go to a ticker you are interested in
    2. Type the function FLDS in your terminal
    3. Search the information you are looking for


# CODE STRUCTURE
I. Five accessible Functions for each type of data request:

        - ReferenceData,
        - PortfolioData,
        - HistoricalData,
        - AIMPortfolioPositionData (UNTESTED),
        - AIMHistPortfolioPositionData (UNTESTED)

They take the inputs and check for errors, convert dates to the Bloomberg format, call catchServerEvent (II) and return the data once it is ready.


II.  The Bloomberg interraction functions:

        - OpenSession
        - OpenService
        - catchServerEvent
        
These private functions handle the connection to Bloomberg's API, the opening of the appropriate service depending on the request and the catching of messages coming from Bloomberg.
catchServerEvent listens to Bloomberg and delegates the processing of messages to the appropriate functions (III).
    

III. Five server data processing functions:

        - getServerData_reference,
        - getServerData_portfolio,
        - getServerData_historical,
        - getServerData_aimportfolio,
        - getServerData_aimhistportfolio
        
These private functions will be called by the catchServerEvent depending on the request sent. Their purpouse is to structure the data for output and catch errors returned by Bloomberg

    
IV. The Data returned: 
is in a matrix of type variant in order to support the different data types returned by Bloomberg
Note that you can usually find descriptive data in the 0th column rows and the 0th row columns, the data itself starts at 1,1

