Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class ForProject

    '**************************************************************************************
    '** 全域變數
    '**************************************************************************************
    Dim Key As String = "WFSDB"     ' 連線字串
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataBaseObj)
    '**     取得資料庫共用函式物件
    '**
    '*****************************************************************
    Public Function GetDataBaseObj() As Utility.DataBase
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(Key))
        Return uDataBase
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Public Function ModifyTranData(ByVal pFun As String, _
                                   ByVal pSts As String, _
                                   ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pNowDateTime As String, _
                                   ByVal pDecideID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQl As String
        SQl = "Update T_WaitHandle Set "
        SQl = SQl + " Active = '" & "0" & "',"
        SQl = SQl + " Sts = '" & pSts & "',"
        SQl = SQl + " StsDes = '" & "完成" & "',"
        SQl = SQl + " AEndTime = '" & pNowDateTime & "',"
        SQl = SQl + " CompletedTime = '" & pNowDateTime & "',"
        SQl = SQl + " DecideDesc = N'" & "OK(自動核定)" & "',"
        SQl = SQl + " ModifyUser = '" & pDecideID & "',"
        SQl = SQl + " ModifyTime = '" & pNowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
        SQl = SQl + "   And Step    =  '" & CStr(pStep) & "'"
        SQl = SQl + "   And SeqNo   =  '" & CStr(pSeqNo) & "'"
        SQl = SQl + "   And Active =  '1' "
        Try
            Dim uDataBase As Utility.DataBase = GetDataBaseObj()
            uDataBase.ExecuteNonQuery(SQl)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Public Function BAModifyTranData(ByVal pFun As String, _
                                   ByVal pSts As String, _
                                   ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pNowDateTime As String, _
                                   ByVal pDecideID As String, _
                                   ByVal pDecideDesc As String, _
                                   ByVal pStsDes As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQl As String
        SQl = "Update T_WaitHandle Set "
        SQl = SQl + " Active = '" & "0" & "',"
        SQl = SQl + " Sts = '" & pSts & "',"
        SQl = SQl + " StsDes = '" & pStsDes & "',"
        SQl = SQl + " AEndTime = '" & pNowDateTime & "',"
        SQl = SQl + " CompletedTime = '" & pNowDateTime & "',"
        SQl = SQl + " DecideDesc = N'" & pDecideDesc & "',"
        SQl = SQl + " ModifyUser = '" & pDecideID & "',"
        SQl = SQl + " ModifyTime = '" & pNowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
        SQl = SQl + "   And Step    =  '" & CStr(pStep) & "'"
        SQl = SQl + "   And SeqNo   =  '" & CStr(pSeqNo) & "'"
        SQl = SQl + "   And Active =  '1' "
        Try
            Dim uDataBase As Utility.DataBase = GetDataBaseObj()
            uDataBase.ExecuteNonQuery(SQl)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Public Function ModifyData(ByVal pFun As String, _
                               ByVal pSts As String, _
                               ByVal pFormNo As String, _
                               ByVal pFormSno As Integer, _
                               ByVal pNowDateTime As String, _
                               ByVal pDecideID As String, _
                               ByVal pTableName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        sql = "Update " & pTableName & " Set "
        sql &= " Sts = '" & pSts & "',"
        sql &= " CompletedTime = '" & pNowDateTime & "',"
        sql &= " ModifyUser = '" & pDecideID & "',"
        sql &= " ModifyTime = '" & pNowDateTime & "' "
        sql &= " Where FormNo  =  '" & pFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(pFormSno) & "'"
        '
        Try
            Dim uDataBase As Utility.DataBase = GetDataBaseObj()
            uDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
End Class
