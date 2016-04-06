
Option Explicit On 
'Option Strict On

Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Collections

Public Class IMCostCentres


#Region "PrivateVariables"

    Private lCostCentreBuildSize As Long
    Private lCostCentreID As Long
    Private strCostCentreName As String
    Private strCostCentreDescription As String
    Private lCostCentreParentID As Long
    Private lCostCentreTypeID As Long 'The cost centre type that the cost centre belongs to
    Private lCostCentreChiefParentTypeID As Long  'The root Cost Centre Type ID
    Private dtDateCreated As Date

#End Region


#Region "Properties"

    Public ReadOnly Property CostCentreBuildSize() As Long

        Get
            Return lCostCentreBuildSize
        End Get


    End Property

    Public Property DateCreated() As Date

        Get
            Return dtDateCreated
        End Get

        Set(ByVal Value As Date)
            dtDateCreated = Value
        End Set

    End Property

    Public Property CostCentreChiefParentTypeID() As Long

        Get
            Return lCostCentreChiefParentTypeID
        End Get

        Set(ByVal Value As Long)
            lCostCentreChiefParentTypeID = Value
        End Set

    End Property

    Public Property CostCentreTypeID() As Long

        Get
            Return lCostCentreTypeID
        End Get

        Set(ByVal Value As Long)
            lCostCentreTypeID = Value
        End Set

    End Property

    Public Property CostCentreParentID() As Long

        Get
            Return lCostCentreParentID
        End Get

        Set(ByVal Value As Long)
            lCostCentreParentID = Value
        End Set

    End Property

    Public Property CostCentreDescription() As String

        Get
            Return strCostCentreDescription
        End Get

        Set(ByVal Value As String)
            strCostCentreDescription = Value
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

    Public Property CostCentreName() As String

        Get
            Return strCostCentreName
        End Get

        Set(ByVal Value As String)
            strCostCentreName = Value
        End Set

    End Property

#End Region


#Region "InitializationProcedures"

    Public Sub New()
        MyBase.New()
    End Sub

#End Region


#Region "GeneralProcedures"

    Public Function ReturnCostCentreNames _
        (Optional ByVal lValCCID As Long = 0) As String()

        Dim strQueryToUse As String
        Dim arProducts() As String
        Dim objLogin As IMLogin

        Try

            If lValCCID = 0 Then
                strQueryToUse = "SELECT CostCentreName FROM CostCentres " & _
                    " ORDER BY CostCentreName ASC"
            Else
                strQueryToUse = "SELECT CostCentreName FROM CostCentres " & _
                    " WHERE CostCentreName = '" & lValCCID & "'"
            End If

            objLogin = New IMLogin

            With objLogin
                arProducts = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arProducts

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnCostCentreID _
        (Optional ByVal strValCCName As String = "") As String()

        Dim strQueryToUse As String
        Dim arProducts() As String
        Dim objLogin As IMLogin

        Try

            If strValCCName = 0 Then
                strQueryToUse = "SELECT CostCentreID FROM CostCentres " & _
                    " ORDER BY CostCentreName ASC"
            Else
                strQueryToUse = "SELECT CostCentreID FROM CostCentres " & _
                    " WHERE CostCentreName = " & strValCCName
            End If

            objLogin = New IMLogin

            With objLogin
                arProducts = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arProducts

        Catch ex As Exception

        End Try

    End Function

    'Done
    Public Function ReturnCostCentreTypeIDFromCostCentreTypeName _
       (ByVal strValCostCentreTypeName As String) As Long()

        Dim strQueryToUse As String
        Dim arCostCentreType() As Long
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT CostCentreTypeID FROM CostCentreTypes " & _
            " WHERE CostCentreTypeName = '" & strValCostCentreTypeName & _
                "' ORDER BY CostCentreTypeName ASC"

            objLogin = New IMLogin

            With objLogin
                arCostCentreType = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arCostCentreType

        Catch ex As Exception

        End Try

    End Function

    'Determines whether the parent selected has the same " & _
    '"Root Parent as the Cost Centre's Parent
    Public Function IsCostCentreCompatibleWithRootAndSelecteParent() _
    As Boolean

        Dim strQueryToUse As String
        Dim arProducts() As String
        Dim objLogin As IMLogin

        Try

            '==========================
            'Check if the selected parent belongs to the same " & _
            '"Cost Centre Type as the provided cost centre

            'first check if the parent selected is acutally a root, i.e. 
            '"it has not lCostCentreParentID
            '+++++++++++++++++++++++++++++++++++++++++
            '+++++++++++++++++++++++++++++++++++++++++
            If Find("SELECT * FROM CostCentres WHERE " & _
            "CostCentreID = " & lCostCentreParentID & _
            " AND CostCentreParentID IS NULL OR CostCentreParentID = 0 ", _
            False) = True Then

                Return True
                objLogin = Nothing

                Exit Function
            End If


            '+++++++++++++++++++++++++++++++++++++++++
            If Find("SELECT * FROM CostCentres WHERE " & _
            "CostCentreID = " & lCostCentreParentID & _
            " AND CostCentreChiefParentTypeID = " & _
            lCostCentreChiefParentTypeID, False) = False Then

                objLogin = Nothing

                Exit Function
            End If

            objLogin = Nothing

            Return True

        Catch ex As Exception

        End Try

    End Function

    Public Function ReturnChiefParentID _
        (ByVal strChiefParentName As String) As Long

        Dim strQueryToUse As String
        Dim arChiefParentID() As Long
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT CostCentreName " & _
                " FROM CostCentres  " & _
                " INNER JOIN CostCentreTypes ON " & _
                " CostCentreTypes.CostCentreTypeID = " & _
                " CostCentres.CostCentreTypeID " & _
                " WHERE CostCentreTypeHierarchyPosition = 1 " & _
                " AND CostCentreName = '" & strChiefParentName & "'"


            objLogin = New IMLogin

            With objLogin
                arChiefParentID = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing


            If Not arChiefParentID Is Nothing Then
                Return arChiefParentID(0)

            End If


        Catch ex As Exception

        End Try

    End Function

    'Array of all Cost Centre Types available
    Public Function ReturnArrayCostCentreTypesAvailable() As String()

        Dim strQueryToUse As String
        Dim arCostCentreTypes() As String
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT CostCentreTypeName FROM CostCentreTypes " & _
                " ORDER BY CostCentreTypeHierarchyPosition ASC"

            objLogin = New IMLogin

            With objLogin
                arCostCentreTypes = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arCostCentreTypes

        Catch ex As Exception

        End Try

    End Function

    'Array of all cost centres that are root cost centres types
    Public Function ReturnArrayCostCentresThatAreParentRoots() As String()

        Dim strQueryToUse As String
        Dim arCostCentreTypes() As String
        Dim objLogin As IMLogin

        Try

            strQueryToUse = "SELECT CostCentreName " & _
                " INNER JOIN CostCentreTypes ON " & _
                " CostCentreTypes.CostCentreTypeID = " & _
                " CostCentres.CostCentreTypeID " & _
                " FROM CostCentres WHERE " & _
                " CostCentreTypeHierarchyPosition = 1 " & _
                " ORDER BY CostCentreName ASC"

            objLogin = New IMLogin

            With objLogin
                arCostCentreTypes = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arCostCentreTypes

        Catch ex As Exception

        End Try


    End Function

    'Get available cost centres of a particular types in a 
    'particular parent root
    Public Function ReturnAvailableCostCentreTypesInATypeInARootNode _
        (ByVal lValCostCentreChiefParentTypeID As String, _
            ByVal strValTypeID As String, _
            Optional ByVal bLimitToOnlyAvailableOnes As Boolean = False) _
                As String()

        Dim strQueryToUse As String
        Dim arProducts() As String
        Dim objLogin As IMLogin
        Dim i As Long


        Try

            strQueryToUse = " SELECT DISTINCT CostCentreName " & _
    " FROM CostCentres " & _
    " INNER JOIN CostCentreTypes ON " & _
    " CostCentreTypes.CostCentreTypeID = " & _
    " CostCentres.CostCentreTypeID " & _
    " WHERE CostCentreChiefParentTypeID = " & _
    lValCostCentreChiefParentTypeID & _
    " AND CostCentreTypeName = '" & strCostCentreName & _
    "' ORDER BY CostCentreName ASC"

            objLogin = New IMLogin

            With objLogin
                arProducts = .FillArray(strOrgAccessConnString, _
                    strQueryToUse, "", "")

            End With

            objLogin = Nothing

            Return arProducts

        Catch ex As Exception

        End Try

    End Function

    Public Function CostCentersArray _
       (ByVal bIncludeCostCenterBranches As Boolean, _
            ByVal lPositionMaxToInclude As Long, _
                ByVal bIncludeEmptyCostCentreBranches As Boolean) _
                    As Object()

        Dim strQueryToUse As String
        Dim arCollectiveCostCentres() As String
        Dim objLogin As IMLogin

        Dim arCostCentreTypeNames() As String
        Dim arParents() As String
        Dim arCostCentreNames() As String

        Dim arResults() As String 'Result is a single string per parent root type node as follows "Type,Parent Name++=Type,Child,Child,Child++=Type,Child,Child,Child

        Dim strItemPosTypes As String 'Will begin with Name then four asterisks i.e ****'
        Dim strItemParents As String
        Dim strItemPosCostCentres As String

        Dim iPosTypes As Long
        Dim iPosParents As Long
        Dim iPosCostCentres As Long


        Try

            'Get the types
            arCostCentreTypeNames = ReturnArrayCostCentreTypesAvailable()

            If arCostCentreTypeNames Is Nothing Then
                ReturnError += "The Cost Centres are empty"
                Exit Function
            End If


            'Get the parents in each type
            arParents = ReturnArrayCostCentresThatAreParentRoots()

            If arParents Is Nothing Then
                ReturnError += "No Cost Centres have been defined"
                Exit Function

            End If

            iPosTypes = 1
            iPosParents = 0
            iPosCostCentres = 0

            'For each parent node
            For Each strItemParents In arParents

                'if the parent nodes are existent
                If strItemParents Is Nothing Then
                    ReturnError += "There are no parent nodes"
                    Exit Function

                    Dim strResultForArray As String
                    Dim strNodeResults As String

                    'Add the Type,Parent Name++= details
                    strResultForArray = arCostCentreTypeNames(iPosParents) & _
                            "," & strItemParents & "++="

                    'For each type, get the children in that parent as array
                    For iPosTypes = 1 To _
                        arCostCentreTypeNames.GetLongLength(0) - 1



                        'Get the cost centre names in this type in this root
                        arCostCentreNames = _
                        ReturnAvailableCostCentreTypesInATypeInARootNode( _
                            ReturnChiefParentID(strItemParents), _
                                arCostCentreTypeNames(iPosTypes))

                        strNodeResults = arCostCentreNames(iPosTypes)

                        '
                        If Not arCostCentreNames Is Nothing Then

                            For Each strItemPosCostCentres In _
                                arCostCentreNames

                                If Not strItemPosCostCentres Is Nothing Then
                                    strNodeResults += "," & _
                                        strItemPosCostCentres

                                End If
                            Next
                        End If

                    Next


                    'Resize array
                    If arResults Is Nothing Then
                        ReDim arResults(1)

                    Else
                        ReDim arResults(arResults.GetLongLength(0) + 1)

                    End If

                    arResults(arResults.GetLongLength(0) - 1) = _
                        strResultForArray & strNodeResults

                    'Seperate the children with bracket and commas
                End If

            Next

            'ReturnArray
            Return arCollectiveCostCentres

        Catch ex As Exception

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

            If Trim(strCostCentreName) = "" Or _
                Trim(strCostCentreDescription) = "" _
                    Or lCostCentreTypeID = 0 Or _
                    lCostCentreChiefParentTypeID = 0 Then

                If DisplayErrorMessages = True Then

                    ReturnError += "Please provide a Cost Centre Title, " & _
                    "its description, and the Cost Centre Type the " & _
                    "Cost Centre belongs to, and the Cost Centre " & _
                    "Root Parent the cost centre belongs to."

                End If

                objLogin = Nothing
                datSaved = Nothing

                Exit Function
            End If


            If IsCostCentreCompatibleWithRootAndSelecteParent() Then

                ReturnError += "The Cost Centre provided must belong " & _
                "to the same Cost Centre Root Parent as the " & _
                "Parent Cost Centre"

                objLogin = Nothing
                datSaved = Nothing

            End If


            If Find("SELECT * FROM CostCentres WHERE " & _
            "CostCentreName = '" & _
            Trim(strCostCentreName) & "'", False) = True Then

                Update("UPDATE CostCentres SET " & _
                            "CostCentreDescription = '" & _
                            Trim(strCostCentreDescription) & _
                            "', CostCentreTypeID = " & _
                            lCostCentreTypeID & _
                            " WHERE CostCentreName = '" & _
                            strCostCentreName & _
                            "' AND CostCentreParentID = " & _
                            lCostCentreParentID & _
                            " AND CostCentreChiefParentTypeID = " & _
                            lCostCentreChiefParentTypeID)


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


            strInsertInto = "INSERT INTO CostCentres (" & _
                "CostCentreName," & _
                "CostCentreDescription," & _
                "lCostCentreParentID," & _
                "CostCentreTypeID," & _
                "CostCentreChiefParentTypeID" & _
                    ") VALUES "

            strSaveQuery = strInsertInto & _
                    "('" & Trim(strCostCentreName) & _
                    "','" & Trim(strCostCentreDescription) & _
                    "'," & CostCentreParentID & _
                    "," & CostCentreTypeID & _
                    "," & CostCentreChiefParentTypeID & _
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

                    ReturnError += "Cost Centre Title Saved Successfully."

                End If

                Return True

            Else

                If DisplayFailure = True Then

                    ReturnError += "'Save Cost Centre Title' action failed." & _
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
                        datRetData = Nothing
                        objLogin = Nothing

                        Return False
                        Exit Function

                    End If

                    'Whether to fill properties with values or not
                    If ReturnStatus = True Then

                        For Each myDataRows In myDataTables.Rows

                            lCostCentreID = _
                                myDataRows("CostCentreID")
                            strCostCentreName = _
                                myDataRows("CostCentreName").ToString
                            strCostCentreDescription = _
                                myDataRows("CostCentreDescription").ToString
                            lCostCentreTypeID = _
                                myDataRows("CostCentreID")
                            lCostCentreParentID = _
                                myDataRows("CostCentreParentID")
                            lCostCentreChiefParentTypeID = _
                                myDataRows("CostCentreChiefParentTypeID")
                            dtDateCreated = _
                               myDataRows("DateCreated")

                        Next
                    End If
                Next

                Return True

            End If



        Catch ex As Exception
            MsgBox(ex.Message.ToString, _
                    MsgBoxStyle.Exclamation, _
                        "iManagement - Critical System Error")

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

            If lCostCentreID = 0 Then
                ReturnError += "Cannot Delete. Please select an " & _
                    "existing Cost Centre Title Detail"

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


            strDeleteQuery = "DELETE * FROM CostCentres WHERE " & _
            "CostCentreID = " & lCostCentreID

            objLogin.ConnectString = strOrgAccessConnString
            objLogin.ConnectToDatabase()

            bDelSuccess = objLogin.ExecuteQuery(strOrgAccessConnString, _
            strDeleteQuery, datDelete)

            objLogin.CloseDb()

            datDelete = Nothing
            objLogin = Nothing

            If bDelSuccess = True Then
                ReturnError += "Cost Centre Title Details Deleted"
                Return True
            Else

                ReturnError += "'Delete Cost Centre action failed"

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

            If lCostCentreID <> 0 Then

                objLogin.ConnectString = strOrgAccessConnString
                objLogin.ConnectToDatabase()

                bUpdateSuccess = objLogin.ExecuteQuery _
                                (strOrgAccessConnString, _
                                    strUpdateQuery, _
                                            datUpdated)

                objLogin.CloseDb()

                If bUpdateSuccess = True Then
                    ReturnError += "Cost Centre Name updated Successfully"
                End If

            End If

            datUpdated = Nothing
            objLogin = Nothing


        Catch ex As Exception

        End Try

    End Sub

#End Region


End Class
