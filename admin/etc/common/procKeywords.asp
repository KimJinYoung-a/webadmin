<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, arrkeywords, cksel, i
Dim vItemid, vKeywords, mallgubun, vidx
mode		= Request("mode")
cksel		= Request("cksel")
arrkeywords	= Request("arrkeywords")
mallgubun	= Request("mallgubun")

If mode = "REG" Then
	If Right(arrkeywords,4) = "*(^!" Then
		arrkeywords = Left(arrkeywords, Len(arrkeywords) - 4)
	End If
	vItemid		= split(cksel, ",")
	vKeywords	= split(arrkeywords, "*(^!")

	For i = LBound(vItemid) to UBound(vItemid)
		sqlStr = ""
		sqlStr = sqlStr & " IF EXISTS (SELECT itemid FROM db_etcmall.[dbo].[tbl_outmall_keywords] WHERE itemid = '"&Trim(vItemid(i))&"' AND mallid = '"& mallgubun &"') " & vbcrlf
		sqlStr = sqlStr & " 	UPDATE db_etcmall.[dbo].[tbl_outmall_keywords] " & vbcrlf
		sqlStr = sqlStr & " 	SET keywords = '"&Trim(vKeywords(i))&"' " & vbcrlf
		sqlStr = sqlStr & " 	, lastupdate = getdate() " & vbcrlf
		sqlStr = sqlStr & " 	WHERE itemid = '"&Trim(vItemid(i))&"' " & vbcrlf
		sqlStr = sqlStr & " 	AND mallid = '"& mallgubun &"' " & vbcrlf
		sqlStr = sqlStr & " ELSE " & vbcrlf
		sqlStr = sqlStr & " 	INSERT INTO db_etcmall.[dbo].[tbl_outmall_keywords] (itemid, keywords, mallid, regdate) " & vbcrlf
		sqlStr = sqlStr & " 	VALUES ('"&Trim(vItemid(i))&"', '"&Trim(vKeywords(i))&"', '"& mallgubun &"', getdate()) "
		dbget.execute sqlStr
	Next
ElseIf mode = "nREG" Then
	vidx		= request("vidx")
	vkeywords	= request("vkeywords")
	sqlStr = ""
	sqlStr = sqlStr & " UPDATE db_etcmall.[dbo].[tbl_outmall_notKeywords] " & vbcrlf
	sqlStr = sqlStr & " SET keywords = '"& vkeywords &"' "
	sqlStr = sqlStr & " WHERE idx = '"&vidx&"' "
	dbget.execute sqlStr
End If
%>
<script language="javascript">
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->