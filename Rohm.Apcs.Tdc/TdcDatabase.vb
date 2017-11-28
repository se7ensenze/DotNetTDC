Imports System.Data
Imports System.Data.SqlClient

Friend Class TdcDatabase

    Private m_ConnectionString As String = My.Settings.APCSDBConnectionString

    Public Property ConnectionString() As String
        Get
            Return m_ConnectionString
        End Get
        Set(ByVal value As String)
            m_ConnectionString = value
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Function CreateConnection() As SqlConnection
        Return New SqlConnection(m_ConnectionString)
    End Function

    Private Function ExecuteDataTable(ByVal connection As SqlConnection, ByVal sqlCommand As String) As DataTable
        Dim dt As DataTable = New DataTable()
        Using cmd As SqlCommand = New SqlCommand(sqlCommand, connection)
            dt.Load(cmd.ExecuteReader())
        End Using
        Return dt
    End Function

    Public Function GetMachine(ByVal connection As SqlConnection, ByVal machineNo As String) As APCSDBDataSet.MACHI_TABLEDataTable
        Dim ret As APCSDBDataSet.MACHI_TABLEDataTable
        Using adaptor As APCSDBDataSetTableAdapters.MACHI_TABLETableAdapter = New APCSDBDataSetTableAdapters.MACHI_TABLETableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByMachineName(machineNo)
        End Using
        Return ret
    End Function

    Public Function GetWorkData(ByVal con As SqlConnection, ByVal lotNo As String) As APCSDBDataSet.WorkDataDataTable
        Dim ret As APCSDBDataSet.WorkDataDataTable
        Using adaptor As APCSDBDataSetTableAdapters.WorkDataTableAdapter = New APCSDBDataSetTableAdapters.WorkDataTableAdapter()
            adaptor.Connection = con
            ret = adaptor.GetDataByLotNo(lotNo)
        End Using
        Return ret
    End Function

    Public Function GetLot1Table(ByVal connection As SqlConnection, ByVal lotNo As String) As APCSDBDataSet.LOT1_TABLEDataTable

        Dim ret As APCSDBDataSet.LOT1_TABLEDataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT1_TABLETableAdapter = New APCSDBDataSetTableAdapters.LOT1_TABLETableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByLotNo(lotNo)
        End Using
        Return ret

    End Function

    Public Function GetLot2Table(ByVal connection As SqlConnection, ByVal lotNo As String) As APCSDBDataSet.LOT2_TABLEDataTable
        Dim ret As APCSDBDataSet.LOT2_TABLEDataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_TABLETableAdapter = New APCSDBDataSetTableAdapters.LOT2_TABLETableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByLotNo(lotNo)
        End Using
        Return ret
    End Function

    Public Function UpdateLot1Table(ByVal con As SqlConnection, ByVal transaction As SqlTransaction, ByVal lot1Table As APCSDBDataSet.LOT1_TABLEDataTable) As Integer

        Dim ret As Integer
        Using adaptor As APCSDBDataSetTableAdapters.LOT1_TABLETableAdapter = New APCSDBDataSetTableAdapters.LOT1_TABLETableAdapter()
            adaptor.Connection = con
            adaptor.Transaction = transaction
            ret = adaptor.Update(lot1Table)
        End Using
        Return ret

    End Function

    Public Function UpdateLot2Table(ByVal con As SqlConnection, ByVal transaction As SqlTransaction, ByVal lot2Table As APCSDBDataSet.LOT2_TABLEDataTable) As Integer
        Dim ret As Integer
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_TABLETableAdapter = New APCSDBDataSetTableAdapters.LOT2_TABLETableAdapter()
            adaptor.Connection = con
            adaptor.Transaction = transaction
            ret = adaptor.Update(lot2Table)
        End Using
        Return ret
    End Function

    Public Function UpdateLot2Data(ByVal con As SqlConnection, ByVal transaction As SqlTransaction, ByVal lot2Data As APCSDBDataSet.LOT2_DATADataTable) As Integer
        Dim ret As Integer
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_DATATableAdapter = New APCSDBDataSetTableAdapters.LOT2_DATATableAdapter()
            adaptor.Connection = con
            adaptor.Transaction = transaction
            ret = adaptor.Update(lot2Data)
        End Using
        Return ret
    End Function

    Public Function GetLot2DataOnLotSet(ByVal connection As SqlConnection, ByVal realStart As Date, ByVal lotNo As String) As APCSDBDataSet.LOT2_DATADataTable
        Dim ret As APCSDBDataSet.LOT2_DATADataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_DATATableAdapter = New APCSDBDataSetTableAdapters.LOT2_DATATableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByPK(lotNo, realStart)
        End Using
        Return ret
    End Function

    Public Function GetLot2DataOnLotEnd(ByVal connection As SqlConnection, ByVal opeSeq As Short, ByVal lotNo As String, _
                                        ByVal mcNo As String, ByVal machineSub As String) As APCSDBDataSet.LOT2_DATADataTable
        Dim ret As APCSDBDataSet.LOT2_DATADataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_DATATableAdapter = New APCSDBDataSetTableAdapters.LOT2_DATATableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByMCNoLotNoOpeSeqMachineSub(opeSeq, lotNo, mcNo, machineSub)
        End Using
        Return ret
    End Function

    Public Function GetLot2DataByLotNoOpeSeq(ByVal connection As SqlConnection, ByVal tran As SqlTransaction, ByVal lotNo As String, ByVal opeSeq As Short) As APCSDBDataSet.LOT2_DATADataTable
        Dim ret As APCSDBDataSet.LOT2_DATADataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_DATATableAdapter = New APCSDBDataSetTableAdapters.LOT2_DATATableAdapter()
            adaptor.Connection = connection
            adaptor.Transaction = tran
            ret = adaptor.GetDataByLotNoOpeSeq(lotNo, opeSeq)
        End Using
        Return ret
    End Function

    Public Function GetLot2DataByLotNoOpeSeq(ByVal connection As SqlConnection, ByVal lotNo As String, ByVal opeSeq As Short) As APCSDBDataSet.LOT2_DATADataTable
        Dim ret As APCSDBDataSet.LOT2_DATADataTable
        Using adaptor As APCSDBDataSetTableAdapters.LOT2_DATATableAdapter = New APCSDBDataSetTableAdapters.LOT2_DATATableAdapter()
            adaptor.Connection = connection
            ret = adaptor.GetDataByLotNoOpeSeq(lotNo, opeSeq)
        End Using
        Return ret
    End Function

    Public Function SumOutputOfLOT2_DATA(ByVal con As SqlConnection, ByVal tran As SqlTransaction, ByVal strLotNo As String, ByVal strLayNo As String, ByRef goodPieces As Integer, ByRef badPieces As Integer) As Boolean

        Dim ret As Boolean = False

        Using cmd As SqlCommand = New SqlCommand()
            cmd.Connection = con
            cmd.Transaction = tran
            cmd.CommandText = "SELECT Sum(B.GOOD_PIECES) AS SUM_GOOD, Sum(B.BAD_PIECES) AS SUM_BAD" & _
            " FROM LOT1_TABLE AS A INNER JOIN LOT2_DATA AS B ON A.LOT_NO = B.LOT_NO" & _
            " WHERE A.LOT_NO ='" & strLotNo & "' AND B.MACHINE_SUB IN (" & CODE_MACHINE_SUB_LOT_END & "," & _
            CODE_MACHINE_SUB_ABNORMAL_ACCUM & ") AND B.LAY_NO ='" & strLayNo & "'"

            Using dt As DataTable = New DataTable()
                dt.Load(cmd.ExecuteReader())
                If dt.Rows.Count = 1 Then
                    Dim row As DataRow = dt.Rows(0)
                    If Convert.IsDBNull(row(0)) Then
                        goodPieces = 0
                    Else
                        goodPieces = CInt(row(0))
                    End If
                    If Convert.IsDBNull(row(1)) Then
                        badPieces = 0
                    Else
                        badPieces = CInt(row(1))
                    End If

                    ret = True

                End If

            End Using
        End Using

        Return ret

    End Function

    Public Function DeleteNotFinishLOT2_DATAOf(ByVal con As SqlConnection, ByVal tran As SqlTransaction, ByVal lotNo As String, ByVal mcNo As String) As Integer
        Dim affectedRow As Integer = 0
        Using cmd As SqlCommand = New SqlCommand()
            cmd.Transaction = tran
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM LOT2_DATA WHERE (LOT_NO = '" & lotNo & "') AND (MACHINE = '" & mcNo & "') AND REAL_DAY IS NULL"
            affectedRow = cmd.ExecuteNonQuery()
        End Using
        Return affectedRow
    End Function

    Public Sub DeleteLot1TableAndLot1DataByLotNo(ByVal con As SqlConnection, ByVal tran As SqlTransaction, ByVal lotNo As String)
        Using cmd As SqlCommand = New SqlCommand()
            cmd.Connection = con
            cmd.Transaction = tran
            cmd.CommandText = "DELETE FROM LOT1_TABLE WHERE LOT_NO = '" & lotNo & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE FROM LOT1_DATA WHERE LOT_NO = '" & lotNo & "'"
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub DeleteLOT2_DATACurrentAndAfterOPE_SEQ(ByVal con As SqlConnection, ByVal tran As SqlTransaction, ByVal lotNo As String, ByVal currentOpeSeq As Short)
        Using cmd As SqlCommand = New SqlCommand()
            cmd.Transaction = tran
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM LOT2_DATA WHERE (LOT_NO = '" & lotNo & "') AND OPE_SEQ >= " & currentOpeSeq.ToString()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Function GetOpeSeqByLotNoLayNo(ByVal con As SqlConnection, ByVal lotNo As String, ByVal layNo As String) As Short
        Dim val As Short = 0
        Using cmd As SqlCommand = New SqlCommand()
            cmd.Connection = con
            cmd.CommandText = "SELECT (OPE_SEQ) FROM LOT1_DATA WHERE (LOT_NO = @LOT_NO) AND (LAY_NO = @LAY_NO)"
            cmd.Parameters.Add("@LOT_NO", SqlDbType.VarChar, 13)
            cmd.Parameters.Add("@LAY_NO", SqlDbType.VarChar, 52)
            cmd.Parameters(0).Value = lotNo
            cmd.Parameters(1).Value = layNo

            Dim obj As Object = cmd.ExecuteScalar()

            If Not Convert.IsDBNull(obj) Then
                val = CShort(obj)
            End If

        End Using
        Return val
    End Function

    Public Function IsFirstProcess(ByVal con As SqlConnection, ByVal lotNo As String, ByVal currentOpeSeq As Short) As Boolean

        Dim ret As Boolean = False

        Using cmd As SqlCommand = New SqlCommand()
            cmd.Connection = con
            cmd.CommandText = "SELECT TOP 1 OPE_SEQ FROM LOT1_DATA WHERE N_OPE_SEQ != 0 AND LOT_NO = @LOT_NO"
            cmd.Parameters.Add("@LOT_NO", SqlDbType.VarChar, 13)
            cmd.Parameters(0).Value = lotNo

            Dim val As Object = cmd.ExecuteScalar()
            If Convert.IsDBNull(val) OrElse val Is Nothing Then
                Throw New Exception("can not find First process")
            Else
                ret = currentOpeSeq.Equals(CType(val, Short))
            End If

        End Using

        Return ret

    End Function

End Class
