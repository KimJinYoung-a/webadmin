<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, categbn, brandcode, makerid, cateCnt
mode = Request("mode")
categbn = Request("categbn")

If (categbn <> "cate" and categbn <> "brand") Then
	response.write "<script>alert('잘못된 경로입니다');window.close();</script>"
	response.end
End If

Dim cdl, cdm, cds, large_category, middle_category, small_category, detail_category
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
large_category	= requestCheckVar(request("large_category"),10)
middle_category	= requestCheckVar(request("middle_category"),10)
small_category	= requestCheckVar(request("small_category"),10)
detail_category	= requestCheckVar(request("detail_category"),10)

If (mode = "saveCate") Then
	If cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
		If detail_category = "0" Then
			sqlStr = ""
			sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
			sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_wetoo1300k_category "
			sqlStr = sqlStr & " WHERE large_category = '"& large_category &"' "
			sqlStr = sqlStr & " and middle_category = '"& middle_category &"' "
			sqlStr = sqlStr & " and small_category = '"& small_category &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
				cateCnt = rsget("cnt")
			rsget.Close

			If cateCnt <> 1 Then
				Call Alert_return("하위카테고리가 존재합니다.\n\n다른 카테고리로 매칭하세요") 
				dbget.close()	:	response.End
			End If
		End If


		sqlStr = "DELETE FROM db_etcmall.dbo.tbl_wetoo1300k_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute sqlStr

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.tbl_wetoo1300k_cate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (large_category, middle_category, small_category, detail_category, tenCateLarge, tenCateMid, tenCateSmall, lastupdate)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & large_category & "', '" & middle_category & "', '" & small_category & "', '" & detail_category & "' "  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
		dbget.execute sqlStr
	Case "delCate"
		sqlStr = "DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] " & VbCrlf
		sqlStr = sqlStr& " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		dbget.execute(sqlStr)
	Case "savebrandcode"
		brandcode	= Trim(request("brandcode"))
		makerid		= Trim(request("makerid"))

		sqlStr = " DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] WHERE makerid ='"& makerid &"' " & VbCrlf
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " INSERT INTO db_etcmall.dbo.[tbl_wetoo1300k_brandcode] (makerid, brandCode, regdate) " & VbCrlf
		sqlStr = sqlStr & " SELECT userid, '"& brandCode &"', GETDATE() " & VbCrlf
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c " & VbCrlf
		sqlStr = sqlStr & " WHERE userid = '"& makerid &"' " 
		dbget.execute(sqlStr)
	Case "delbrandcode"
		makerid		= Trim(request("makerid"))
		brandcode	= Trim(request("brandcode"))
		sqlStr = " DELETE FROM db_etcmall.dbo.[tbl_wetoo1300k_brandcode] WHERE makerid ='"& makerid &"' and brandcode = '"& brandcode &"' " & VbCrlf
		dbget.execute(sqlStr)
End Select
%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
	alert("<%=iErrMsg %>");
	history.back(-1);
<% Else %>
	alert("정상적으로 처리되었습니다.");
	<% If mode = "savebrandcode" or mode = "delbrandcode" Then %>
	location.replace('/admin/etc/wetoo1300k/popwetoo1300kBrandList.asp?brandcode=<%= brandCode %>');
	<% Else %>
	parent.opener.history.go(0);
	parent.self.close();
	<% End If %>
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->