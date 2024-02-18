<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
	'// 저장 모드 접수
	dim mode, sqlStr
	mode = Request("mode")

    '// 상품번호/옵션번호를 받는다 //
    dim dispNo, cdl, cdm, cds
    dispNo = Request("dspNo")
    cdl = Request("cdl")
    cdm = Request("cdm")
    cds = Request("cds")

	if dispNo="" or cdl="" or cdm="" or cds="" then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	end if

	'// 모드별 분기
	Select Case mode
		Case "save"
			'중복 확인
			sqlStr = "Select DispNo From db_item.dbo.tbl_lotte_cate_mapping " &_
					" Where tenCateLarge='" & cdl & "'" &_
					" 	and tenCateMid='" & cdm & "'" &_
					" 	and tenCateSmall='" & cds & "'" &_
					" 	and DispNo='" & dispNo & "'"
			rsget.Open sqlStr,dbget,1
			if rsget.EOF then
				'신규등록
				sqlStr = "Insert into db_item.dbo.tbl_lotte_cate_mapping values " &_
						" ('" & dispNo & "'" &_
						", '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
				dbget.execute(sqlStr)
			end if
			rsget.Close

		Case "del"
			'매칭된 텐바이텐 카테고리 삭제
			sqlStr = "Delete From db_item.dbo.tbl_lotte_cate_mapping " &_
					" Where tenCateLarge='" & cdl & "'" &_
					" 	and tenCateMid='" & cdm & "'" &_
					" 	and tenCateSmall='" & cds & "'" &_
					" 	and DispNo='" & dispNo & "'"
			dbget.execute(sqlStr)
	End Select
	
	if (mode="save") or (mode="del") then
	    CALL Fn_ActOutMall_CateSummary("lotteCom")
	end if
%>
<script language="javascript">
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->