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
Dim mode, sqlStr, iErrMsg, idx, i
mode = Request("mode")
idx = request("idx")

'// ��庰 �б�
Select Case mode
	Case "saveCate"
		For i = 1 to request("catecode").count
			If LEN(request("catecode")(i)) < 9 Then
				response.write "<script>alert('3Depth������ ī�װ��� �����ֽ��ϴ�.');history.back(-1);</script>"
				response.end
			End If

			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(*) as cnt FROM db_outmall.[dbo].[tbl_coocha_cate_mapping] WHERE tencatecode = '"&LEFT(request("catecode")(i), 9)&"' and depthCode <> '"&idx&"' "
			rsCTget.Open sqlStr,dbCTget,1
			If rsCTget("cnt") > 0 Then
				response.write "<script>alert('ī�װ� �߿� �̹� ��ϵ� ī�װ��� �����ֽ��ϴ�.');history.back(-1);</script>"
				response.end
			End If
			rsCTget.Close
		Next

		sqlStr = " DELETE FROM db_outmall.[dbo].[tbl_coocha_cate_mapping] WHERE depthCode = '"&idx&"' "
		dbCTget.execute(sqlStr)
		For i = 1 to request("catecode").count
			sqlStr = " IF NOT Exists(SELECT * FROM db_outmall.[dbo].[tbl_coocha_cate_mapping] WHERE depthCode = '"&idx&"' and tenCateCode = '"&LEFT(request("catecode")(i), 9)&"' ) "
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " 	INSERT INTO db_outmall.[dbo].[tbl_coocha_cate_mapping] "
	        sqlStr = sqlStr & " 	(depthCode, tenCateCode, lastupdate)"
	        sqlStr = sqlStr & " 	VALUES ("&idx&", '"&LEFT(request("catecode")(i), 9)&"' ,getdate())"
			sqlStr = sqlStr & " END "
			dbCTget.execute(sqlStr)
		Next

		sqlStr = ""
		sqlStr = sqlStr & " UPDATE db_outmall.[dbo].[tbl_coocha_cate_mapping] SET "
		sqlStr = sqlStr & " DEPTH1NM = B.DEPTH1NM "
		sqlStr = sqlStr & " ,DEPTH2NM = B.DEPTH2NM "
		sqlStr = sqlStr & " ,DEPTH3NM = B.DEPTH3NM "
		sqlStr = sqlStr & " FROM db_outmall.[dbo].[tbl_coocha_cate_mapping] A "
		sqlStr = sqlStr & " JOIN db_outmall.[dbo].[tbl_coocha_category] B on A.depthCode = B.idx "
		sqlStr = sqlStr & " WHERE A.depthCode = '"&idx&"' "
		dbCTget.execute(sqlStr)

	Case "delCate"
		'��Ī�� �ٹ����� ī�װ� ����
		sqlStr = "Delete From db_outmall.[dbo].[tbl_coocha_cate_mapping] " & VbCrlf
		sqlStr = sqlStr& " Where depthCode = '"&idx&"' "
		dbCTget.execute(sqlStr)
End Select
%>
<script language="javascript">
alert("���������� ó���Ǿ����ϴ�.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->