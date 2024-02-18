<%
Function WaitItemCheckMyItemYN(MakerID, ItemID)
	Dim sqlStr
	sqlstr = "select top 1 itemid from db_academy.dbo.tbl_diy_wait_item"
	sqlstr = sqlstr & " where itemid='" & ItemID & "'"
	sqlstr = sqlstr & " and makerid='" & MakerID & "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	If Not rsACADEMYget.Eof Then
		WaitItemCheckMyItemYN = True
	Else
		WaitItemCheckMyItemYN = False
	End If
	rsACADEMYget.Close
End Function

Function ItemCheckMyItemYN(MakerID, ItemID)
	Dim sqlStr
	sqlstr = "select top 1 itemid from db_academy.dbo.tbl_diy_item"
	sqlstr = sqlstr & " where itemid='" & ItemID & "'"
	sqlstr = sqlstr & " and makerid='" & MakerID & "'"
	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	If Not rsACADEMYget.Eof Then
		ItemCheckMyItemYN = True
	Else
		ItemCheckMyItemYN = False
	End If
	rsACADEMYget.Close
End Function
%>