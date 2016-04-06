
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMObjectCoordinates

#Region "PrivateVariables"

    Private lCoordinateID As Long
    Private strCoordinateType As Long 'Similar to strFixedAssetsIMCategory
    Private dbXCoordinate As Double
    Private dbYCoordinate As Double
    Private dbDegree As Double
    Private lObjectKeyID As Long
    Private lCostCentreID As Long
    Private strEstateName As String
  

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

    Public Property CoordinateID() As Long

        Get
            Return lCoordinateID
        End Get

        Set(ByVal Value As Long)
            lCoordinateID = Value
        End Set

    End Property

    Public Property CoordinateType() As String

        Get
            Return strCoordinateType
        End Get

        Set(ByVal Value As String)
            strCoordinateType = Value
        End Set

    End Property

    Public Property XCoordinate() As Double

        Get
            Return dbXCoordinate
        End Get

        Set(ByVal Value As Double)
            dbXCoordinate = Value
        End Set

    End Property

    Public Property YCoordinate() As Double

        Get
            Return dbYCoordinate
        End Get

        Set(ByVal Value As Double)
            dbYCoordinate = Value
        End Set

    End Property

    Public Property Degree() As Double

        Get
            Return dbDegree
        End Get

        Set(ByVal Value As Double)
            dbDegree = Value
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

    Public Property CostCentreID() As Long

        Get
            Return lCostCentreID
        End Get

        Set(ByVal Value As Long)
            lCostCentreID = Value
        End Set

    End Property

    Public Property EstateName() As String

        Get
            Return strEstateName
        End Get

        Set(ByVal Value As String)
            strEstateName = Value
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

            If (dbYCoordinate = 0 And dbXCoordinate <> 0) Or _
            (dbYCoordinate <> 0 And dbXCoordinate = 0) Then

                ReturnError += "Invalid details entered. Please enter " & _
                    "both the X and the Y coordinates"

                objLogin = Nothing
                datSaved = Nothing

                Exit Function

            End If


            If Trim(strCoordinateType) = "" Or lObjectKeyID = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide the following details in" & _
                " order to save the Object's Coordinates: " & _
                Chr(10) & "1. Coordinate Type " & _
                Chr(10) & "2. Object to who's coordinate you want to save "
                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If

            If Find("SELECT * FROM ObjectCoordinates " & _
    " WHERE (CoordinateType = '" & strCoordinateType & _
    "' AND XCoordinate = " & dbXCoordinate & _
    " AND YCoordinate = " & dbYCoordinate & ")", _
    False) = True Then

                If Find("SELECT * FROM ObjectCoordinateOverlapping " & _
        " WHERE (CoordinateType = '" & strCoordinateType & _
        "' AND OverlappingAllowed = TRUE)", _
        False) = False Then

                    ReturnError += "This particular Object's Coordinate " & _
                        "Type does not allow overlapping of coordinates"

                    objLogin = Nothing
                    datSaved = Nothing

                    Exit Function
                End If

            End If


            If Find("SELECT * FROM ObjectCoordinates " & _
                " WHERE (ObjectKeyID = " & lObjectKeyID & _
                " AND CoordinateType = '" & strCoordinateType & "')", _
                False) = False Then

                ReturnError += "The Object already exists."

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            strInsertInto = "INSERT INTO ObjectCoordinates (" & _
                "CoordinateID," & _
                "CoordinateType," & _
                "XCoordinate," & _
                "YCoordinate," & _
                "Degree," & _
                "ObjectKeyID," & _
                "CostCentreID," & _
                "EstateName" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "(" & lCoordinateID & _
                    ",'" & strCoordinateType & _
                    "'," & dbXCoordinate & _
                    "," & dbYCoordinate & _
                    "," & dbDegree & _
                    "," & lObjectKeyID & _
                    "," & lCostCentreID & _
                    ",'" & strEstateName & _
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
                    ReturnSuccess += "Coordinate Saved Successfully."

                End If
            Else

                If DisplayFailure = True Then
                    ReturnError = "'Save Coorinate' action failed." & _
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

                            lCoordinateID = _
                                myDataRows("CoordinateID")
                            strCoordinateType = _
                                myDataRows("CoordinateType")
                            dbXCoordinate = _
                                myDataRows("XCoordinate")
                            dbYCoordinate = _
                                myDataRows("YCoordinate")
                            dbDegree = _
                                myDataRows("Degree")
                            lObjectKeyID = _
                                myDataRows("ObjectKeyID")
                            lCostCentreID = _
                                myDataRows("CostCentreID")
                            strEstateName = _
                                myDataRows("EstateName")

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

            strDeleteQuery = "DELETE * FROM ObjectCoordinates " & _
            "WHERE ObjectKeyID = " & lObjectKeyID

            If lObjectKeyID = 0 Then

                ReturnError += "You must provide an object with a coordinate " & _
                        "in order to delete it"
                Exit Function
            End If

            objLogin.ConnectString = strAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
                    strDeleteQuery, datDelete)

            objLogin.CloseDb()

            If bDelSuccess = True Then
                ReturnSuccess += "Fixed Asset details deleted"

            Else

                ReturnError += "'Delete Fixed Asset' action failed"


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
                        ReturnSuccess += "Object's Coordinate record updated successfully"
                    End If

                End If

            End If

        Catch ex As Exception

        End Try
    End Sub


#End Region


#Region "GeneralProcedures"

    Public Function ReturnObjectsWithinARadiusOfXInTheseCoordinates _
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

    Public Function ReturnObjectCoordinateObject _
        (ByVal lValObjectKeyID As Long, _
            ByVal strValCoordinateType As String) As Object


    End Function

    Public Function ReturnObjectsWithinCostCentre _
        (ByVal lValCostCentre As Long, _
            ByVal strValCoordinateType As String) As Object


    End Function

    Public Function ReturnObjectsOfThisTypeLinkedToProvidedCoordinateObject _
        (ByVal strValCoordinateType As String) As Object


    End Function



#End Region

End Class
