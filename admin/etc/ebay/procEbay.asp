<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// ���� ��� ����
Dim mode, sqlStr, iErrMsg, categbn, makerid, BrandCode
mode    = Request("mode")
categbn = Request("categbn")


Dim cdl, cdm, cds, cateCode, stdcode, i
Dim chkArr, cateCodeArr, stdcodeArr, gubunArr
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
cateCode = requestCheckvar(Request("cateCode"),30)
stdcode   = requestCheckvar(Request("stdcode"),30)

If (mode = "saveCate") or (mode="saveCateArr") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("���۵� ���� �����ϴ�.\nó���� ����Ǿ����ϴ�.","about:blank")
		dbget.Close: response.End
	End If
End If

'// ��庰 �б�
Select Case mode
    CASE "saveCateArr"
        set chkArr = Request.form("chk")
        set stdcodeArr = Request.form("stdcode")
        set cateCodeArr= Request.form("cateCode")
        set gubunArr  = Request.form("gubun")

        If chkArr.count <> "2" Then
            Call Alert_move("2���� ī�װ��� �������ּ���.\nó���� ����Ǿ����ϴ�.","about:blank")
            dbget.Close: response.End
        End IF

        If stdcodeArr(chkArr(1)+1) <> stdcodeArr(chkArr(2)+1) Then
            Call Alert_move("���� ESMī�װ� �������� �������ּ���.\nó���� ����Ǿ����ϴ�.","about:blank")
            dbget.Close: response.End
        End IF

        If gubunArr(chkArr(1)+1) = gubunArr(chkArr(2)+1) Then
            Call Alert_move("���ǰ� ������ �ϳ��� ���� ���ּ���.\nó���� ����Ǿ����ϴ�.","about:blank")
            dbget.Close: response.End
        End IF

		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_etcmall.[dbo].[tbl_ebay_cate_mapping] " & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)

        For i=1 To chkArr.count
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_ebay_cate_mapping] " & VbCrlf
			sqlStr = sqlStr & " (SDCategoryCode, cateCode, tenCateLarge, tenCateMid, tenCateSmall, gubun, lastupdate) VALUES " & VbCrlf
			sqlStr = sqlStr & " ('"& stdcodeArr(chkArr(i)+1) &"', '"& cateCodeArr(chkArr(i)+1) &"', '"& cdl &"', '"& cdm &"', '"& cds &"', '"& gubunArr(chkArr(i)+1) &"',getdate()) "
			dbget.execute(sqlStr)
        Next

        set chkArr = Nothing
        set stdcodeArr = Nothing
        set cateCodeArr = Nothing
        set gubunArr = Nothing
	Case "delCate"
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_etcmall.[dbo].[tbl_ebay_cate_mapping] " & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
    alert("���������� ó���Ǿ����ϴ�.");

    opener.location.reload();
    parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->