<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim vQuery, vMakerID, vTemp, i, j, vDispCateTmp, standardCateCode
vMakerID = requestCheckVar(request("makerid"),100)
vTemp = request("filecnt")
standardCateCode = request("standardCateCode")

If standardCateCode <> "" Then
	vQuery =""
	vQuery = vQuery & " UPDATE db_user.dbo.tbl_user_c SET standardCateCode = '"&standardCateCode&"' WHERE userid = '"&vMakerID&"' "
	dbget.execute vQuery
End If

vQuery = ""

For i = 1 To vTemp
	If request("dispcate"&i&"") <> "" Then
		vDispCateTmp = vDispCateTmp & request("dispcate"&i&"") & ","
	End If
Next

Dim vSameCount, tmp, vIsOverLap, vIsStandardOverLap
vIsOverLap = "x"
For i = LBound(Split(vDispCateTmp,",")) To UBound(Split(vDispCateTmp,","))-1
	tmp = Split(vDispCateTmp,",")(i)
	
	If tmp = standardCateCode Then
		vIsStandardOverLap = "o"
	End If
	
	For j = LBound(Split(vDispCateTmp,",")) To UBound(Split(vDispCateTmp,","))-1
		If tmp = Split(vDispCateTmp,",")(j) Then
			vSameCount = vSameCount + 1
		End If
		
		If vSameCount > 1 Then
			vIsOverLap = "o"
		End If
	Next
	vSameCount = 0
Next

If vIsStandardOverLap = "o" Then
	Response.Write "<script>alert('��ǥ ����ī�װ��� �ߺ��� �ڵ带 �����ϼ̽��ϴ�.\n�ߺ� ���ؼ� �������ּ���.');history.back();</script>"
	dbget.close()
	Response.End
ElseIf vIsOverLap = "o" Then
	Response.Write "<script>alert('�ߺ��� �ڵ带 �����ϼ̽��ϴ�.\n�ϳ��� �������ּ���.');history.back();</script>"
	dbget.close()
	Response.End
Else
	For i = 1 To vTemp
		If request("dispcate"&i&"") <> "" Then
			vQuery = vQuery & " INSERT INTO [db_partner].[dbo].[tbl_partner_dispcate](makerid, catecode) VALUES('" & vMakerID & "', '" & request("dispcate"&i&"") & "') " & vbCrLf
		End If
	Next
End If

If standardCateCode <> "" Then
	vQuery = vQuery & " INSERT INTO [db_partner].[dbo].[tbl_partner_dispcate](makerid, catecode, isdefault) VALUES('" & vMakerID & "', '" &standardCateCode& "', 'Y') " & vbCrLf
End If

vQuery = "DELETE [db_partner].[dbo].[tbl_partner_dispcate] WHERE makerid = '" & vMakerID & "' " & vbCrLf & vQuery
dbget.execute vQuery
%>
<script>
location.href = "/admin/member/popbrandinfoonly_dispcate.asp?makerid=<%=vMakerID%>";
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->