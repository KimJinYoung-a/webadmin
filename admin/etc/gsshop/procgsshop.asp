<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg, categbn
mode = Request("mode")
'// 상품번호/옵션번호를 받는다 //
Dim dispNo, cdl, cdm, cds, safecode
dispNo	= requestCheckvar(Request("dspNo"),32)
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
safecode	= requestCheckvar(Request("safecode"),10)
categbn = requestCheckvar(Request("categbn"),1)

If (mode = "saveCate") OR (mode = "delGbn") OR (mode = "delCate") OR (mode = "delNewPrddiv") OR (mode = "saveNewPrdDiv") Then
	If (dispNo = "" ) OR cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
		'매칭된 전시구분 삭제
		sqlStr = ""
		sqlStr = sqlStr & " Delete From db_item.dbo.tbl_gsshop_cate_mapping " & VbCrlf
		sqlStr = sqlStr & " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr & " and categbn ='" & categbn & "'"
		dbget.execute(sqlStr)

		sqlStr = ""
		sqlStr = sqlStr & " Insert into db_item.dbo.tbl_gsshop_cate_mapping  " & VbCrlf
		sqlStr = sqlStr & " (CateKey, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate, categbn)" & VbCrlf
		sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
		sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate(), '"& categbn &"') "
		dbget.execute sqlStr
	Case "delCate"
		'매칭된 텐바이텐 카테고리 삭제
		sqlStr = "Delete From db_item.dbo.tbl_gsshop_cate_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
		dbget.execute(sqlStr)

	Case "saveNewPrdDiv"
        '중복 확인
        sqlStr = ""
		sqlStr = sqlStr & " SELECT dtlCd FROM db_item.dbo.tbl_gsshop_MngDiv_mapping "  & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'"  & VbCrlf
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget.EOF Then
			'신규등록
			sqlStr = ""
			sqlStr = sqlStr & " INSERT INTO db_item.dbo.tbl_gsshop_MngDiv_mapping  " & VbCrlf
			sqlStr = sqlStr & " (dtlCd, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate, safecode)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '" & cdl & "','" & cdm & "','" & cds & "', getdate(),'" & safecode & "') "
			dbget.execute sqlStr
		Else
		    iErrMsg = "이미 매핑된 상품분류는 ["&dispNo&"] 추가할 수 없습니다."
		End If
		rsget.Close

	Case "delNewPrddiv"
		'매칭된 텐바이텐 카테고리 삭제
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_item.dbo.tbl_gsshop_MngDiv_mapping " & VbCrlf
		sqlStr = sqlStr & " WHERE tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr & " and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr & " and dtlCd = '" & dispNo & "'"
		dbget.execute(sqlStr)
End Select

If (mode="saveCate") or (mode="delCate") then
    CALL Fn_ActOutMall_CateSummary("gsshop")
End If
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