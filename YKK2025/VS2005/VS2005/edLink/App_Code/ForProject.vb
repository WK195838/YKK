Imports Microsoft.VisualBasic
'---------------------------------------------------------------------------------------------------
'***************************************************************************************************
'** Database
'***************************************************************************************************
'Database-Start
Imports System.Data             'SQL
Imports System.Data.SqlClient
'Database-End
'---------------------------------------------------------------------------------------------------

Public Class ForProject

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '** 全域變數
    '***********************************************************************************************
    '全域變數-Start
    Dim Key As String = "EDIDB"             ' 連線字串
    '全域變數-End

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**(GetDataBaseObj)
    '**     取得資料庫共用函式物件
    '***********************************************************************************************
    'GetDataBaseObj-Start
    Public Function GetDataBaseObj() As Utility.DataBase
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(Key))
        Return uDataBase
    End Function
    'GetDataBaseObj-End

    '-----------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([600]-AccessLog)
    '**     解析結果寫入記錄檔
    '** LogID, Buyer, Action, RunTime, Error, 說明1, 說明2, 說明3, 說明4, 說明5, UserID, UserName
    '***********************************************************************************************
    'AccessLog-Start
    Public Sub AccessLog(ByVal pLogID As String, ByVal pBuyer As String, _
                         ByVal pAction As String, ByVal pRunTime As String, _
                         ByVal pError As Integer, _
                         ByVal pDescr1 As String, ByVal pDescr2 As String, _
                         ByVal pDescr3 As String, ByVal pDescr4 As String, _
                         ByVal pDescr5 As String, _
                         ByVal pUserID As String, ByVal pUserName As String)
        Dim SQL As String
        Dim oDataBase As Utility.DataBase = GetDataBaseObj()
        '
        SQL = "Insert into L_ActionHistory "
        SQL &= "( "
        SQL &= "LogID, Buyer, Action, RunTime, Error, "
        SQL &= "D1, D2, D3, D4, D5, "
        SQL &= "UserID, UserName, CreateUser, CreateTime "
        SQL &= ")  "
        SQL &= "VALUES( "
        SQL &= " '" & pLogID & "', "
        SQL &= " '" & pBuyer & "', "
        SQL &= " '" & pAction & "', "
        SQL &= " '" & pRunTime & "', "
        SQL &= " '" & CStr(pError) & "', "
        SQL &= " '" & pDescr1 & "', "
        SQL &= " '" & pDescr2 & "', "
        SQL &= " '" & pDescr3 & "', "
        SQL &= " '" & pDescr4 & "', "
        SQL &= " '" & pDescr5 & "', "
        SQL &= " '" & pUserID & "', "
        SQL &= " '" & pUserName & "', "
        SQL &= " '" & pAction & "', "
        SQL &= " '" & pRunTime & "' "
        SQL &= " ) "
        oDataBase.ExecuteNonQuery(SQL)
    End Sub
    'AccessLog-End

    '---------------------------------------------------------------------------------------------------
    '***********************************************************************************************
    '**([610]-GetFunctionCode)
    '**     取得Buyer Group中Function Code
    '***********************************************************************************************
    'GetFunctionCode-Start
    Public Function GetFunctionCode(ByVal pBuyerGroup As String, ByVal pFun As Integer) As String
        Dim RtnString As String = ""
        '
        Try
            RtnString = Mid(pBuyerGroup, pFun, 1)
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    'GetFunctionCode-End


End Class
