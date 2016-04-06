
Module basDatabase

    'Version in use
    Public bIsDemoVersion As Boolean
    Public bDesktopInterface As Boolean
    Public bWebInterface As Boolean

    'Business Process Work Flow
    Public bCanEditWorkFlow As Boolean

    'Work Flow Types
    Public bHasDocumentManagement As Boolean
    Public bHasBillingManagament As Boolean
    Public bHasCRM As Boolean
    Public bHasHRMini As Boolean

    'HR
    Public bHasHRMedium As Boolean
    Public hHasHRFull As Boolean

    'Fixed assets
    Public bHasFixedAssets As Boolean


    'Buildings
    Public bHasBuildings As Boolean

    'Appartments
    Public bHasAppartments As Boolean


    Public strDBUserName As String
    Public strDBPassword As String
    Public strDBDatabase As String
    Public strDBDBPath As String
    Public bDBLoggedIn As Boolean

    Public strAccessConnString As String
    Public strSQLConnString As String
    Public bConnectionSQLServer As Boolean
    Public strAccessMdw As String
    Public strAccessConnStringADOX As String

    'Transaction As Boolean
    Public bCommadTransactionInitiate As Boolean
    Public bCommadTransactionStartedState As Boolean
    Public bCommandTransactionCompleteState As Boolean


    Public strOrgAccessConnString As String
    Public strOrgSQLConnString As String
    Public bOrgConnectionSQLServer As Boolean
    Public strOrgAccessMdw As String
    Public strOrgAccessConnStringADOX As String

    Public strOrganizationName As String
    Public strOrganizationID As String
    Public strOrgDBPath As String
    Public MainFormHeading As String

    Public strCurrentForm As String
    Public strCurrentSearch As String
    Public strCurrentSearchFor As String
    Public bUseDefaultFilters As Boolean

    Public ReturnError As String
    Public ReturnSuccess As String

    'Quotation Mark (Chr(34))
    Public strQuote


End Module
