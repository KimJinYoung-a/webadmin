<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, id, sourcearea, sqlStr
mode		= Request("mode")
id			= Request("sid")
sourcearea	= Request("sname")

If mode = "I" Then
	sqlStr = ""
	sqlStr = sqlStr & " IF EXISTS (SELECT id FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] WHERE id = '"&Trim(id)&"' AND sourcearea = '"& Trim(sourcearea) &"') " & vbcrlf
	sqlStr = sqlStr & " 	UPDATE db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] " & vbcrlf
	sqlStr = sqlStr & " 	SET id = '"&Trim(id)&"' " & vbcrlf
	sqlStr = sqlStr & " 	, sourcearea = getdate() " & vbcrlf
	sqlStr = sqlStr & " 	WHERE id = '"& Trim(sourcearea) &"' " & vbcrlf
	sqlStr = sqlStr & " 	AND sourcearea = '"& Trim(sourcearea) &"' " & vbcrlf
	sqlStr = sqlStr & " ELSE " & vbcrlf
	sqlStr = sqlStr & " 	INSERT INTO db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] (id, sourcearea) " & vbcrlf
	sqlStr = sqlStr & " 	VALUES ('"&Trim(id)&"', '"& Trim(sourcearea) &"') "
	dbget.execute sqlStr
ElseIf mode = "D" Then
	sqlStr = ""
	sqlStr = sqlStr & " DELETE FROM db_etcmall.[dbo].[tbl_ssg_sourceAreaCodeMapping] " & vbcrlf
	sqlStr = sqlStr & " WHERE id = '"&Trim(id)&"' AND sourcearea = '"& sourcearea &"' " & vbcrlf
	dbget.execute sqlStr
End If
%>
<script language="javascript">
parent.location.reload();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->