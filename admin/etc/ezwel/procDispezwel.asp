<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, joongBok, categbn, tDepthCode
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "cate")  Then
	response.write "<script>alert('�߸��� ����Դϴ�');window.close();</script>"
	response.end
End If

Dim catecode, depthCode
catecode	= Request("catecode")
depthCode	= requestCheckvar(Request("depthcode"),10)
joongBok = False
If (mode = "saveCate") Then
	If catecode = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
        '�ߺ� Ȯ��
		sqlStr = "Delete From db_etcmall.dbo.tbl_ezwel_dispcate_mapping Where catecode= '"& catecode &"' " & VbCrlf
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_ezwel_dispcate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (depthCode, catecode, lastUpdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
		sqlStr = sqlStr & ", '"& catecode &"', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "Delete From db_etcmall.dbo.tbl_ezwel_dispcate_mapping Where catecode= '"& catecode &"' " & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->