<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, joongBok, categbn
mode = Request("mode")
categbn = Request("categbn")
If (categbn <> "cate")  Then
	response.write "<script>alert('�߸��� ����Դϴ�');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, depthCode
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
depthCode= requestCheckvar(Request("depthcode"),10)
joongBok = False
If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

Dim tDepthCode
'// ��庰 �б�
Select Case mode
	Case "saveCate"
        '�ߺ� Ȯ��
        If categbn = "cate" Then
	        sqlStr = "select top 1 C.depthCode from "
			sqlStr = sqlStr& " db_etcmall.[dbo].[tbl_auction_category_New] as C "
			sqlStr = sqlStr& " JOIN db_etcmall.dbo.tbl_auction_cate_mapping as m on c.depthcode = m.depthcode "
			sqlStr = sqlStr& " WHERE tenCateLarge = '" & cdl & "' and tenCateMid = '" & cdm & "' and tenCateSmall = '" & cds & "'  "
			rsget.Open sqlStr,dbget,1
			If not rsget.EOF Then
				tDepthCode = rsget("depthCode")
			End If
			rsget.Close

			If tDepthCode = "" Then
				sqlStr = "Delete From db_etcmall.dbo.tbl_auction_cate_mapping " & VbCrlf
				sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
				sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
				sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
				dbget.execute(sqlStr)
			End If

	        sqlStr = "SELECT COUNT(*) as cnt FROM db_etcmall.dbo.tbl_auction_cate_mapping "  & VbCrlf
			sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
'			sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
			rsget.Open sqlStr,dbget,1
			If rsget("cnt") > 0 Then
			     joongBok = True
			End If
			rsget.Close
		End If

		If joongBok = False Then
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_auction_cate_mapping  " & VbCrlf
			sqlStr = sqlStr & " (depthCode, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & depthCode & "' "  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbget.execute sqlStr
		Else
		    iErrMsg = "�̹� ���ε� ī�װ���  �߰��� �� �����ϴ�."
		End If
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_auction_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and depthCode='" & depthCode & "'"
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