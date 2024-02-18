<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, joongBok, categbn
mode	= Request("mode")
categbn	= Request("categbn")

Dim cdl, cdm, cds
Dim std_cat_id, disp_cat_id
cdl			= requestCheckvar(Request("cdl"),10)
cdm			= requestCheckvar(Request("cdm"),10)
cds			= requestCheckvar(Request("cds"),10)
std_cat_id	= requestCheckvar(Request("std_cat_id"), 30)
disp_cat_id	= requestCheckvar(Request("disp_cat_id"), 30)
joongBok = False
If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        '중복 확인
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_etcmall.dbo.tbl_lotteon_cate_mapping " & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.[dbo].[tbl_lotteon_cate_mapping] " & VbCrlf
		sqlStr = sqlStr & " (std_cat_id, disp_cat_id, tenCateLarge, tenCateMid, tenCateSmall, lastupdate) VALUES " & VbCrlf
		sqlStr = sqlStr & " ('"& std_cat_id &"', '"& disp_cat_id &"', '"& cdl &"', '"& cdm &"', '"& cds &"', getdate()) "
		dbget.execute(sqlStr)
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.[dbo].[tbl_lotteon_cate_mapping] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->