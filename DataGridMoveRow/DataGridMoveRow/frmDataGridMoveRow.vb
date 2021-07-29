#Region "ABOUT"
' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More: http://www.g2gnet.com/webboard
' /
' / Purpose: Move up and down selected rows in DataGridView.
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
#End Region

Public Class frmDataGridMoveRow

    Private Sub frmDataGridMoveRow_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Call InitDataGrid()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Initialized DataGridView and put the sample data.
    Private Sub InitDataGrid()
        '// Initialize DataGridView Control
        With DataGridView1
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ReadOnly = True
            .Font = New Font("Tahoma", 9)
            ' Autosize Column
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '.AutoResizeColumns()
            '// Even-Odd Color
            .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue
            ' Adjust Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.Black ' Color.White
                .Font = New Font("Tahoma", 9, FontStyle.Bold)
            End With
        End With
        '// Declare columns type.
        Dim Column1 As New DataGridViewTextBoxColumn()
        Dim Column2 As New DataGridViewTextBoxColumn()
        Dim Column3 As New DataGridViewTextBoxColumn()
        '// Add new Columns
        DataGridView1.Columns.AddRange(New DataGridViewColumn() { _
                Column1, Column2, Column3 _
                })
        With DataGridView1
            .Columns(0).Name = "Product ID"
            .Columns(1).Name = "Product Name"
            .Columns(2).Name = "Unit Price"
        End With

        '// SAMPLE DATA
        Dim RandomClass As New Random()
        For iCount As Byte = 1 To 10
            Dim row = New String() { _
            iCount, "Product " & iCount, Format(RandomClass.Next(999) + RandomClass.NextDouble(), "0.00")}
            DataGridView1.Rows.Add(row)
        Next
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Move Up.
    Private Sub btnUp_Click(sender As System.Object, e As System.EventArgs) Handles btnUp.Click
        With Me.DataGridView1
            '// หาค่า Index แถวที่เลือก
            Dim RowIndex As Integer = .SelectedCells(0).OwningRow.Index
            '// หาก Index = 0 แสดงว่าเป็นแถวบนสุด ให้จบออกจากโปรแกรมย่อย
            If RowIndex = 0 Then Return
            '//
            Dim Col As Integer = .SelectedCells(0).OwningColumn.Index
            Dim Rows As DataGridViewRowCollection = .Rows
            '// เก็บค่าแถวที่เลือก
            Dim Row As DataGridViewRow = Rows(RowIndex)
            '// ลบแถวที่เลือกออก
            Rows.Remove(Row)
            '// ไปเพิ่มแถวใหม่ ก่อนแถวที่เลือก 1 แถว (ก็เลยเสมือนมันเคลื่อนย้ายแถวได้)
            Rows.Insert(RowIndex - 1, Row)
            '// เคลียร์การเลือกแถว
            .ClearSelection()
            '// โฟกัสรายการแถวที่เลื่อนขึ้นไปแทรก
            .Rows(RowIndex - 1).Cells(Col).Selected = True
        End With
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Move Down.
    Private Sub btnDown_Click(sender As System.Object, e As System.EventArgs) Handles btnDown.Click
        With Me.DataGridView1
            Dim RowIndex As Integer = .SelectedCells(0).OwningRow.Index
            '// เช็คแถวสุดท้าย หากใช่ให้รีเทิร์นกลับ คือไม่ต้องทำอะไร
            If RowIndex = .Rows.Count - 1 Then Return
            '//
            Dim Col As Integer = .SelectedCells(0).OwningColumn.Index
            Dim Rows As DataGridViewRowCollection = .Rows
            '// เก็บค่าแถวที่เลือก
            Dim Row As DataGridViewRow = Rows(RowIndex)
            '// ลบแถวที่เลือกออก
            Rows.Remove(Row)
            '// ไปเพิ่มแถวใหม่ หลังแถวที่เลือก 1 แถว
            Rows.Insert(RowIndex + 1, Row)
            '// เคลียร์การเลือกแถว
            .ClearSelection()
            '// โฟกัสรายการแถวที่เลื่อนลงไป
            .Rows(RowIndex + 1).Cells(Col).Selected = True
        End With
    End Sub

    '// การจัดเรียงตัวเลขให้ถูกต้องด้วยการเปรียบเทียบค่า (Compare)
    Private Sub DataGridView1_SortCompare(sender As Object, e As System.Windows.Forms.DataGridViewSortCompareEventArgs) Handles DataGridView1.SortCompare
        Select Case e.Column.Name
            Case "Product ID"
                e.SortResult = CInt(e.CellValue1).CompareTo(CInt(e.CellValue2))
                e.Handled = True
            Case "Product Name"
                e.SortResult = CStr(e.CellValue1).CompareTo(CStr(e.CellValue2))
                e.Handled = True
            Case "Unit Price"
                e.SortResult = CDbl(e.CellValue1).CompareTo(CDbl(e.CellValue2))
                e.Handled = True
        End Select
    End Sub
End Class
