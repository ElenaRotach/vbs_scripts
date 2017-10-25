Sub Uppercase()
    For Each Cell In oExcel.Selection
        If Not Cell.HasFormula Then
			if Cell.Value <> "" then
				Cell.Value = UCase(Cell.Value)
			end if
            
        End If
    Next
End Sub

sub main
   oExcel.Range("A1:AX300").Select
   Uppercase
   msgbox("end")
end sub
	Set oExcel = GetObject(, "Excel.Application")
	Set Sheet = oExcel.ActiveWorkBook.WorkSheets(1)
	call main