
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMBuildingMaster
    Inherits IMBuildingCostCentre

#Region "PrivateVariables"

    Private strBuildingName As String
    Private strPhysicalAddress As String
    Private strCountry As String
    Private strCity As String
    Private strTown As String
    Private strStreet1 As String
    Private strStreet2 As String
    Private strStreet3 As String
    Private strStreet4 As String
    Private strCurrentHouseNumber As String
    Private lBuildingTypeID As Long
    Private strBuildingHousingNumber As String
    Private strBuildingDescription As String
    Private lNoOfFloors As Long
    Private strRentType As String
    Private dtDateCreated As Date
    Private strBuildSNO As String
    Private strBldColor As String

#End Region


#Region "Properties"

    Public Property BuildingSNo() As String

        Get
            Return strBuildSNO
        End Get

        Set(ByVal Value As String)
            strBuildSNO = Value
        End Set

    End Property


    Public Property BuildingName() As String

        Get
            Return strBuildingName
        End Get

        Set(ByVal Value As String)
            strBuildingName = Value
        End Set

    End Property


    Public Property PhysicalAddress() As String

        Get
            Return strPhysicalAddress
        End Get

        Set(ByVal Value As String)
            strPhysicalAddress = Value
        End Set

    End Property


    Public Property Street1() As String

        Get
            Return strStreet1
        End Get

        Set(ByVal Value As String)
            strStreet1 = Value
        End Set

    End Property


    Public Property Street2() As String

        Get
            Return strStreet2
        End Get

        Set(ByVal Value As String)
            strStreet2 = Value
        End Set

    End Property


    Public Property Street3() As String

        Get
            Return strStreet3
        End Get

        Set(ByVal Value As String)
            strStreet3 = Value
        End Set

    End Property

    Public Property Street4() As String

        Get
            Return strStreet4
        End Get

        Set(ByVal Value As String)
            strStreet4 = Value
        End Set

    End Property

    Public Property BuildingHousingNumber() As String

        Get
            Return strBuildingHousingNumber
        End Get

        Set(ByVal Value As String)
            strBuildingHousingNumber = Value
        End Set

    End Property

    Public Property CurrentHouseNumber() As String

        Get
            Return strCurrentHouseNumber
        End Get

        Set(ByVal Value As String)
            strCurrentHouseNumber = Value
        End Set

    End Property

    Public Property BuildingTypeID() As Long

        Get
            Return lBuildingTypeID
        End Get

        Set(ByVal Value As Long)
            lBuildingTypeID = Value
        End Set

    End Property

    Public Property BuildingDescription() As String

        Get
            Return strBuildingDescription
        End Get

        Set(ByVal Value As String)
            strBuildingDescription = Value
        End Set

    End Property

    Public Property NoOfFloors() As Long

        Get
            Return lNoOfFloors
        End Get

        Set(ByVal Value As Long)
            lNoOfFloors = Value
        End Set

    End Property

    Public Property RentType() As String

        Get
            Return strRentType
        End Get

        Set(ByVal Value As String)
            strRentType = Value
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

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    '[Gets the next invoice number
    Private Function CalculateNextBuildingSerialNo() As Long

        Try
            Dim lReturnValue As Long

            Dim objLogin As IMLogin = New IMLogin

            With objLogin
                lReturnValue = _
                    .ReturnMaxLongValue(strOrgAccessConnString, _
                    "SELECT Count(*) FROM BuildingMaster")

            End With


            objLogin = Nothing

            Return Microsoft.VisualBasic.DateAndTime.Day(Now()) _
                    & Month(Now()) & _
                        Year(Now()) & CostCentreID & "-" & lReturnValue + 1

        Catch ex As Exception
            ReturnError += ex.Message
        End Try

    End Function



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

            If CostCentreID = 0 Then
                ReturnError += "Please provide a valid cost centre"

                datSaved = Nothing
                objLogin = Nothing
                Exit Function
            End If

            If Find("SELECT * FROM BuildingMaster " & _
            "WHERE BuildSNO = " & strBuildSNO, False) _
            = True Then

                Update("UPDATE Building SET " & _
                " BuildingName = '" & strBuildingName & _
                "',PhysicalAddress = '" & strPhysicalAddress & _
                "',Street1 = '" & strStreet1 & _
                "',Street2 = '" & strStreet2 & _
                "',Street3 = '" & strStreet3 & _
                "',Street4 = '" & strStreet4 & _
                "',BuildingHousingNumber = '" & strBuildingHousingNumber & _
                "',CurrentHouseNumber = '" & strCurrentHouseNumber & _
                "',BuildingType = " & lBuildingTypeID & _
                ",BuildingDescription = '" & strBuildingDescription & _
                "',NoOfFloors = " & lNoOfFloors & _
                ",RentType = '" & strRentType & _
                "',BuildingColor = '" & strBldColor & "'")

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            'If MsgBox(" Are you sure you want to add this Cost Centre?", _
            '            MsgBoxStyle.YesNo, _
            '            "iManagement - Add New Record?") = _
            '            MsgBoxResult.No Then

            '    datSaved = Nothing
            '    objLogin = Nothing
            '    Exit Function
            'End If



            strInsertInto = "INSERT INTO BuildingMaster (" & _
                "BuildingName," & _
                "PhysicalAddress," & _
                "Street1," & _
                "Street2," & _
                "Street3," & _
                "Street4," & _
                "BuildingHousingNumber," & _
                "CurrentHouseNumber," & _
                "BuildingTypeID," & _
                "BuildingDescription," & _
                "NoOfFloors," & _
                "RentType," & _
                "BuildingSerialNo," & _
                "BldColor" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & strBuildingName & _
                    "','" & strPhysicalAddress & _
                    "','" & strStreet1 & _
                    "','" & strStreet2 & _
                    "','" & strStreet3 & _
                    "','" & strStreet4 & _
                    "','" & strBuildingHousingNumber & _
                    "','" & strCurrentHouseNumber & _
                    "'," & lBuildingTypeID & _
                    ",'" & strBuildingDescription & _
                    "'," & lNoOfFloors & _
                    ",'" & strRentType & _
                    "'," & CalculateNextBuildingSerialNo() & _
                    ",'" & strRentType & _
                            "')"

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bSaveSuccess = objLogin.ExecuteQuery _
                (strOrgAccessConnString, _
            strSaveQuery, _
            datSaved)


            objLogin.CloseDb()

            If bSaveSuccess = True Then
                If DisplaySuccess = True Then

                    ReturnError += "Cost Centre Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Cost Centre' action failed." & _
            " Make sure all mandatory details are entered."

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If

            objLogin = Nothing
            datSaved = Nothing


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
                    (strOrgAccessConnString, strQuery, datRetData)

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
                        datRetData = Nothing
                        objLogin = Nothing

                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows
                            BuildingID = _
                                myDataRows("BuildingID")
                            strBuildingName = _
                                myDataRows("BuildingName")
                            strPhysicalAddress = _
                                myDataRows("PhysicalAdddress")
                            strStreet1 = _
                                myDataRows("Street1")
                            strStreet2 = _
                                myDataRows("Street2")
                            strStreet3 = _
                                myDataRows("Street4")
                            strStreet4 = _
                                myDataRows("Street3")
                            strBuildingHousingNumber = _
                                myDataRows("BuildingHousingNumber")
                            strCurrentHouseNumber = _
                               myDataRows("CurrentHouseNumber")

                            lBuildingTypeID = _
                                myDataRows("BuildingTypeID")
                            strBuildingDescription = _
                                myDataRows("BuildingDescription")
                            lNoOfFloors = _
                                myDataRows("NoOfFloors")
                            strRentType = _
                                myDataRows("RentType")
                            dtDateCreated = _
                                myDataRows("DateCreated")
                            strBuildSNO = _
                               myDataRows("BuildSNO")
                            strBldColor = _
                                myDataRows("BldColor")

                        Next
                    End If
                Next

                Return True

            End If



        Catch ex As Exception
            ReturnError += ex.Message.ToString

        End Try

    End Function

    Public Function Delete(Optional ByVal bDisplayError As Boolean = False, _
        Optional ByVal bDisplayConfirm As Boolean = False, _
            Optional ByVal bDisplaySuccess As Boolean = False) As Boolean

        Try

            Dim strDeleteQuery As String
            Dim datDelete As DataSet = New DataSet
            Dim bDelSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            If BuildingID = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Building item"

                datDelete = Nothing
                objLogin = Nothing

                Exit Function
            End If

            'If bDisplayConfirm = True Then
            '    If MsgBox(" Are you sure you want to Delete this Cost Centre?", _
            '                             MsgBoxStyle.YesNo, _
            '                             "iManagement - Delete Record?") = _
            '                              MsgBoxResult.No Then

            '        datDelete = Nothing
            '        objLogin = Nothing
            '        Exit Function
            '    End If
            'End If


            strDeleteQuery = "DELETE * FROM BuildingMaster WHERE " & _
            "BuildingID = " & BuildingID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Building's Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Building' action failed"

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Sub Update(ByVal strUpQuery As String)

        Try

            Dim strUpdateQuery As String
            Dim datUpdated As DataSet = New DataSet
            Dim bUpdateSuccess As Boolean
            Dim objLogin As IMLogin = New IMLogin

            strUpdateQuery = strUpQuery

            If BuildingID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Building's details updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region



End Class
