Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMObjectMeterReadings

#Region "PrivateVariables"

    Private lRecordID As Long
    Private dtReadingDate As Date
    Private dbAmountRead As Double
    Private dtDateCreated As Date
    Private strObjectIMCategoryType As String 'Same as strFixedAssetIMCategory
    Private lObjectKeyID As Long

#End Region

#Region "Properties"

    Public Property ReturnError() As Long

        Get
            Return ReturnError
        End Get

        Set(ByVal Value As Long)
            ReturnError = Value
        End Set

    End Property

    Public Property RecordID() As Long

        Get
            Return lRecordID
        End Get

        Set(ByVal Value As Long)
            lRecordID = Value
        End Set

    End Property

    Public Property ReadingDate() As Date

        Get
            Return dtReadingDate
        End Get

        Set(ByVal Value As Date)
            dtReadingDate = Value
        End Set

    End Property

    Public Property AmountRead() As Double

        Get
            Return dbAmountRead
        End Get

        Set(ByVal Value As Double)
            dbAmountRead = Value
        End Set

    End Property

    Public Property DateCreated() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property ObjectIMCaegoryType() As String

        Get
            Return strObjectIMCategoryType
        End Get

        Set(ByVal Value As String)
            strObjectIMCategoryType = Value
        End Set

    End Property

    Public Property ObjectKeyID() As Long

        Get
            Return lObjectKeyID
        End Get

        Set(ByVal Value As Long)
            lObjectKeyID = Value
        End Set

    End Property


#End Region

#Region "InitializationProcedures"

        Public Sub New()
            MyBase.New()
        End Sub

#End Region

#Region "DatabaseProcedures"

        Public Function Save(ByVal DisplayErrorMessages As Boolean, _
            ByVal DisplayConfirmation As Boolean, _
                ByVal DisplayFailure As Boolean, _
                    ByVal DisplaySuccess As Boolean) As Boolean

            Dim strSaveQuery As String
            Dim datSaved As DataSet = New DataSet
            Dim bSaveSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin
            Dim strInsertInto As String

            Try

            If dbAmountRead = 0 Or _
                Trim(strObjectIMCategoryType) = "" Or _
            lObjectKeyID = 0 Then
                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                Chr(10) & " order to save the Object's Meter Readings: " & _
                Chr(10) & "1. Existing Object " & _
                Chr(10) & "1. Object's Meter Reading Amount " & _
                Chr(10) & "2. Object's Type "
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'Check if the object exists as an asset
            If Find("SELECT * FROM FixedAssets " & _
       " WHERE (FixedAssetID = " & lObjectKeyID & _
       " AND FixedAssetIMCategory = '" & _
       Trim(strObjectIMCategoryType) & "')", _
       False) = False Then

                ReturnError += "This particular Object's Fixed Asset details do not exist"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            'If the object's details have been added
            If Find("SELECT * FROM ObjectMeterReadings " & _
    " WHERE (ReadingDate = #" & dtReadingDate & _
    "# AND ObjectIMCategoryType = '" & strObjectIMCategoryType & _
    "' AND ObjectKeyID = " & ObjectKeyID & ")", _
    False) = True Then

            
                ReturnError += "This particular Object's Meter" & _
                    " Reading Details for the date provided already" & _
                "exist"

                    objLogin = Nothing
                    datSaved = Nothing

            End If


            strInsertInto = "INSERT INTO ObjectMeterReadings (" & _
                "RecordID," & _
                "ReadingDate," & _
                "AmountRead," & _
                "DateCreated," & _
                "ObjectIMCategoryType," & _
                "ObjectKeyID," & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lRecordID & _
                    ",#" & dtReadingDate & _
                    "#," & dbAmountRead & _
                    ",#" & dtDateCreated & _
                    "#,'" & strObjectIMCategoryType & _
                    "'," & lObjectKeyID & _
                            ")"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then
                    ReturnError += "Meter readings Saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Meter Readings' action failed." & _
            " Make sure all mandatory details are entered"

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing

            Return True

        Catch ex As Exception
                If DisplayErrorMessages = True Then
                    ReturnError += ex.Message.ToString

                End If
            End Try

        End Function

        Public Function Find(ByVal strQuery As String, _
                            ByVal ReturnStatus As Boolean) As Boolean
            'Query must contain at least rows from Sequence

            Try

                Dim datRetData As DataSet = New DataSet
                Dim bQuerySuccess As Boolean
                Dim myDataTables As DataTable
                Dim myDataColumns As DataColumn
                Dim myDataRows As DataRow
                Dim objLogin As IMLogin = New IMLogin

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bQuerySuccess = objLogin.ExecuteQuery _
                        (strOrgAccessConnString, strQuery, _
                                                        datRetData)

                objLogin.CloseDb()

                If datRetData Is Nothing Then
                    Exit Function
                End If

            If bQuerySuccess = True Then

                Dim i As Integer

                For Each myDataTables In datRetData.Tables

                    'Check if there is any data. If not exit
                    If myDataTables.Rows.Count = 0 Then

                        'Return a value indicating that the search was not successful
                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lRecordID = _
                                myDataRows("RecordID")
                            dtReadingDate = _
                                myDataRows("ReadingDate")
                            dbAmountRead = _
                                myDataRows("AmountRead")
                            dtDateCreated = _
                                myDataRows("DateCreated")
                            strObjectIMCategoryType = _
                                myDataRows("ObjectIMCategoryType")
                            lObjectKeyID = _
                                myDataRows("ObjectKeyID")

                        Next

                    End If

                Next
                Return True

            Else
                Return False

            End If

        Catch ex As Exception
                ReturnError += ex.Message.ToString

            End Try

        End Function

    Public Function Delete() As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strDeleteQuery = "DELETE * FROM ObjectMeterReadings " & _
            "WHERE ObjectKeyID = " & lObjectKeyID

            If lObjectKeyID = 0 Then

                ReturnError += "You must provide an object and its " & _
                    "accompanying meter reading date"
            End If

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnError += "Meter reading details deleted"

            Else

                ReturnError += "'Delete meter reading' action failed"


            End If


        Catch ex As Exception

        End Try

    End Function

    Public Sub Update(ByVal strUpQuery As String, _
    ByVal DisplayErrorMessages As Boolean, _
        ByVal DisplayConfirmation As Boolean, _
            ByVal DisplayFailure As Boolean, _
                ByVal DisplaySuccess As Boolean)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If lObjectKeyID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                                    strUpdateQuery, datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    If DisplaySuccess = True Then
                        ReturnError += "Object's Meter readings record updated successfully"
                    End If

                End If

            End If

        Catch ex As Exception

        End Try
    End Sub


#End Region

#Region "GeneralProcedures"

    Public Function ReturnMeterReadingsForObjectWithinThisRadius _
        (ByVal dbValXCoordinate As Double, _
            ByVal dbValYCoordinate As Double, _
                ByVal strUnitName As String, _
                    ByVal dbRadiusLength As Double, _
                        Optional ByVal strValCoordinateType As String = "") _
                            As Object

        Try



        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnObjectMeterReadings _
        (ByVal lValObjectKeyID As Long, _
            ByVal strValObjectIMCategoryType As String, _
                 ByVal dtStartDate As Date, _
                    ByVal dtEndDate As Date, _
                        ByVal bIncludeDates As Boolean) As Object

        Try


        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnObjectsMeterReadingsWithinCostCentre _
        (ByVal lValCostCentre As Long, _
            ByVal strValCoordinateType As String) As Object

        Try


        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnObjectMeterReadingsForObjectOfThisTypeLinkedToProvidedCoordinateObject _
        (ByVal strValCoordinateType As String) As Object

        Try


        Catch ex As Exception

        End Try

    End Function

#End Region


End Class
