<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim i, vAction, vQuery, vTotalCnt, vItemID, vItemOption, vTplcode, vPcode, vTplname
	vAction		= Request("gubun")
	vTotalCnt	= Request("itemid").count
	vItemID		= Split(Request("itemid"),",")
	vItemOption	= Split(Request("itemoption"),",")
	vTplcode	= Split(Request("tplcode"),",")
	vPcode		= Split(Request("pcode"),",")
	vTplname	= Split(Request("tplname"),",")
	
	vQuery = ""

	For i = 0 To vTotalCnt-1
		If vItemID(i) <> "" Then
			vQuery = vQuery & " IF EXISTS(SELECT itemoption from [db_item].[dbo].[tbl_fuji_templete_code] where itemid = '" & Trim(vItemID(i)) & "' and itemoption = '" & Trim(vItemOption(i)) & "') " & _
							  "		BEGIN " & _
							  "			UPDATE [db_item].[dbo].[tbl_fuji_templete_code] " & _
							  "				SET tplcode = '" & Trim(vTplcode(i)) & "', pcode = '" & Trim(vPcode(i)) & "', tplname = '" & Trim(vTplname(i)) & "' " & _
							  "			WHERE itemid = '" & Trim(vItemID(i)) & "' AND itemoption = '" & Trim(vItemOption(i)) & "' " & _
							  "		END " & _
							  "	ELSE " & _
							  "		BEGIN " & _
							  "			INSERT INTO [db_item].[dbo].[tbl_fuji_templete_code](itemid, itemoption, tplcode, pcode, tplname) " & _
							  "			VALUES('" & Trim(vItemID(i)) & "','" & Trim(vItemOption(i)) & "','" & Trim(vTplcode(i)) & "','" & Trim(vPcode(i)) & "','" & Trim(vTplname(i)) & "') " & _
							  "		END "
		End IF
	Next

	If vQuery <> "" Then
		dbget.execute vQuery
	End If
	
	Response.Write "<Script>alert('저장되었습니다.');location.href='/admin/itemmaster/pop_photobook.asp';</script>"
	Response.End
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->