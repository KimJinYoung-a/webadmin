<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
	'// 저장 모드 접수
	dim mode, sqlStr, iErrMsg
	mode = Request("mode")

    '// 상품번호/옵션번호를 받는다 //
    dim dispNo, cdl, cdm, cds, itemGbnKey '', odispNo, oitemGbnKey
    dispNo  = requestCheckvar(Request("dspNo"),32)
    ''odispNo = requestCheckvar(Request("odspNo"),32)
    ''itemGbnKey = requestCheckvar(Request("itemGbnKey"),32)
    ''oitemGbnKey = requestCheckvar(Request("oitemGbnKey"),32)
    
    cdl = requestCheckvar(Request("cdl"),10)
    cdm = requestCheckvar(Request("cdm"),10)
    cds = requestCheckvar(Request("cds"),10)

    if (mode="saveCate") or (mode="delGbn") or (mode="delCate") then
    	if (dispNo="" ) or cdl="" or cdm="" or cds=""  then
    		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
    		dbget.Close: response.End
    	end if
    end if
    
	'// 모드별 분기
	Select Case mode
		Case "saveCate"
''			'중복 확인 //사용안함
''			sqlStr = "Select cateKey From db_item.dbo.tbl_LTiMall_cateGbn_mapping "  & VbCrlf
''			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
''			sqlStr = sqlStr& " 	and cateKey='" & oitemGbnKey & "'"
''			rsget.Open sqlStr,dbget,1
''			if rsget.EOF then
''				'신규등록
''				sqlStr = "Insert into db_item.dbo.tbl_LTiMall_cateGbn_mapping  "  & VbCrlf
''				sqlStr = sqlStr& " (tenCateLarge,tenCateMid,tenCateSmall,CateKey,lastUpdate)"
''				sqlStr = sqlStr& " values('" & cdl & "','" & cdm & "','" & cds & "','" & itemGbnKey & "', getdate()) "
''				dbget.execute(sqlStr)
''			else
''			    '업데이트
''			    sqlStr = "update db_item.dbo.tbl_LTiMall_cateGbn_mapping  "  & VbCrlf
''			    sqlStr = sqlStr& " set cateKey='"&itemGbnKey&"'"
''				sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
''    			sqlStr = sqlStr& " 	and cateKey='" & oitemGbnKey & "'"
''				dbget.execute(sqlStr)
''			end if
''			rsget.Close
            
            '중복 확인
            sqlStr = "Select cateKey From db_item.dbo.tbl_LTiMall_cate_mapping "  & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
			rsget.Open sqlStr,dbget,1

			if rsget.EOF then
				'신규등록
				sqlStr = "Insert into db_item.dbo.tbl_LTiMall_cate_mapping  "  & VbCrlf
				sqlStr = sqlStr& " (CateKey,tenCateLarge,tenCateMid,tenCateSmall,lastUpdate)"
				sqlStr = sqlStr& " values('" & dispNo & "'"  & VbCrlf
				sqlStr = sqlStr& ", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute sqlStr
			else
			    iErrMsg = "이미 매핑된 카테고리 ["&dispNo&"] 추가할 수 없습니다."
			end if
			rsget.Close

		Case "delCate"
			'매칭된 텐바이텐 카테고리 삭제
			sqlStr = "Delete From db_item.dbo.tbl_LTiMall_cate_mapping " & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & dispNo & "'"
			dbget.execute(sqlStr)
	    Case "delGbn"
	        '매칭된 텐바이텐 상품분류 삭제
			sqlStr = "Delete From db_item.dbo.tbl_LTiMall_cateGbn_mapping " & VbCrlf
			sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
			sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
			sqlStr = sqlStr& " 	and cateKey='" & itemGbnKey & "'"
			dbget.execute(sqlStr)
	End Select
	
	if (mode="saveCate") or (mode="delCate") then
	    CALL Fn_ActOutMall_CateSummary("lotteimall")
	end if
%>
<script language="javascript">
<% if (iErrMsg<>"") then %>
alert("<%=iErrMsg %>");
<% else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->