<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, categbn
mode = Request("mode")
categbn = Request("categbn")

If (categbn <> "cate") Then
	response.write "<script>alert('�߸��� ����Դϴ�');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, lgroup, mgroup, sgroup, dgroup
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
lgroup	= requestCheckVar(request("lgroup"),10)
mgroup	= requestCheckVar(request("mgroup"),10)
sgroup	= requestCheckVar(request("sgroup"),10)
dgroup	= requestCheckVar(request("dgroup"),10)

If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_skstoa_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_skstoa_cate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (lgroup, mgroup, sgroup, dgroup, tenCateLarge, tenCateMid, tenCateSmall, lastupdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & lgroup & "', '" & mgroup & "', '" & sgroup & "', '" & dgroup & "' "  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_skstoa_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
	alert("<%=iErrMsg %>");
	history.back(-1);
<% Else %>
	alert("���������� ó���Ǿ����ϴ�.");
	parent.opener.history.go(0);
	parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->