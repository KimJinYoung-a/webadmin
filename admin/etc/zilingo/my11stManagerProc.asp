<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->
<%
Dim vQuery, vItemID, vItemName, vRegUserID, i, sqlstr
vItemID = Request("itemid")
vRegUserID = session("ssBctId")

If vItemID = "" Then
	Response.Write "<script>alert('잘못된 경로입니다.');window.close()</script>"
	dbget.close()
	Response.End
End IF
vItemName = chrbyte(html2db(Request("itemname")),60,"")

sqlstr = ""
sqlstr = sqlstr & " IF EXISTS( " & VBCRLF
sqlstr = sqlstr & "		SELECT TOP 1 * " & VBCRLF
sqlstr = sqlstr & "		FROM db_etcmall.[dbo].[tbl_my11st_regItem] " & VBCRLF
sqlstr = sqlstr & "		WHERE itemid = '" & vItemID & "' " & VBCRLF
sqlstr = sqlstr & "	)" & VBCRLF
sqlstr = sqlstr & "		UPDATE db_etcmall.[dbo].[tbl_my11st_regItem] SET " & VBCRLF
sqlstr = sqlstr & "		transItemname = '"& vItemName &"'" & VBCRLF
sqlstr = sqlstr & "		WHERE itemid = '" & vItemID & "' " & VBCRLF
sqlstr = sqlstr & " ELSE" & VBCRLF
sqlstr = sqlstr & " 	INSERT INTO db_etcmall.[dbo].[tbl_my11st_regItem] (" & VBCRLF
sqlstr = sqlstr & " 	itemid, reguserid, transItemname " & VBCRLF
sqlstr = sqlstr & "		) VALUES (" & VBCRLF
sqlstr = sqlstr & "		'" & vItemID & "', '" & vRegUserID & "', '"&vItemName&"' " & VBCRLF
sqlstr = sqlstr & "		)"

'response.write sqlstr &"<Br>"
dbget.execute sqlstr

'### 옵션저장.
Dim vOptionCount, vItemOption, vOptionTypeName, vOptionName, vOptIsUsing
vOptionCount = chrbyte(Request("optioncount"),3,"")

vQuery = ""
For i=0 To vOptionCount-1
	vItemOption 	= Request("itemoption"&i)
	vOptionTypeName	= html2db(Request("optiontypename"&i))
	vOptionName		= html2db(Request("optionname"&i))
	vOptIsUsing		= Request("optisusing"&i)

	If vItemOption<>"0000" then
		vQuery = vQuery & "IF EXISTS(SELECT itemoption FROM db_etcmall.[dbo].[tbl_my11st_option] WHERE itemid = '" & vItemID & "' AND itemoption = '" & vItemOption & "') " & vbCrLf
		vQuery = vQuery & "BEGIN " & vbCrLf
		vQuery = vQuery & "		UPDATE db_etcmall.[dbo].[tbl_my11st_option] SET" & vbCrLf
		vQuery = vQuery & "			optiontypename = '" & vOptionTypeName & "', " & vbCrLf
		vQuery = vQuery & "			optionname = '" & vOptionName & "', " & vbCrLf
		vQuery = vQuery & "			isusing = N'" & vOptIsUsing & "' " & vbCrLf
		vQuery = vQuery & "		WHERE itemid = '" & vItemID & "' AND itemoption = '" & vItemOption & "'" & vbCrLf
		vQuery = vQuery & "END " & vbCrLf
		vQuery = vQuery & "ELSE " & vbCrLf
		vQuery = vQuery & "BEGIN " & vbCrLf
		vQuery = vQuery & "		INSERT INTO db_etcmall.[dbo].[tbl_my11st_option](itemid, itemoption, optiontypename, optionname, isusing) " & vbCrLf
		vQuery = vQuery & "		VALUES(N'" & vItemID & "', '" & vItemOption & "', '" & vOptionTypeName & "', '" & vOptionName & "', '" & vOptIsUsing & "') " & vbCrLf
		vQuery = vQuery & "END " & vbCrLf

		vItemOption		= ""
		vOptionTypeName	= ""
		vOptionName		= ""
		vOptIsUsing		= ""
	End If
Next
'rw vQuery
If vQuery <> "" Then
	dbget.execute vQuery
End If
%>

<script type="text/javascript">
	alert("저장되었습니다.");
	opener.document.location.reload();
	window.close();
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->