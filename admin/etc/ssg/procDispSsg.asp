<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, joongBok, categbn, siteNo
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "cate")  Then
	response.write "<script>alert('�߸��� ����Դϴ�');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, DispCtgId
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
DispCtgId= requestCheckvar(Request("DispCtgId"),20)
siteNo	= requestCheckvar(Request("siteNo"),4)
joongBok = False
If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
	Case "saveCate"
        ' '�ߺ� Ȯ��
        ' If categbn = "cate" Then
	    '     sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] "  & VbCrlf
		' 	sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		' 	sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
		' 	rsget.Open sqlStr,dbget,1
		' 	If rsget("cnt") > 0 Then
		' 	     joongBok = True
		' 	End If
		' 	rsget.Close
		' End If

		' If joongBok = False Then
		' 	sqlStr = ""
		' 	sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_DispCate_mapping]  " & VbCrlf
		' 	sqlStr = sqlStr & " (dispCtgId, tenCateLarge, tenCateMid, tenCateSmall, siteNo, lastUpdate)" & VbCrlf
		' 	sqlStr = sqlStr & " VALUES('" & DispCtgId & "' "  & VbCrlf
		' 	sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', '"&siteNo&"', getdate()) "
		' 	dbget.execute sqlStr
		' Else
		'     iErrMsg = "�̹� ���ε� ī�װ���  �߰��� �� �����ϴ�."
		' End If
		sqlStr = "DELETE FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_ssg_DispCate_mapping]  " & VbCrlf
		sqlStr = sqlStr & " (dispCtgId, tenCateLarge, tenCateMid, tenCateSmall, siteNo, lastUpdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & DispCtgId & "' "  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', '"&siteNo&"', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.[dbo].[tbl_ssg_DispCate_mapping] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and dispCtgId='" & DispCtgId & "'"
		sqlStr = sqlStr& " 	and siteNo='" & siteNo & "'"
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
<!-- #include virtual="/lib/db/dbclose.asp" -->