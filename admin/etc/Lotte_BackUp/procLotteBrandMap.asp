<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
	'// 저장 모드 접수
	dim mode, sqlStr
	mode = Request("mode")

    '// 상품번호/옵션번호를 받는다 //
    dim TenMakerid, lotteBrandCd, lotteBrandNm
    TenMakerid = Request("TenMakerid")
    lotteBrandCd = Request("lotteBrandCd")
    lotteBrandNm = Request("lotteBrandNm")

	if TenMakerid="" or lotteBrandCd="" then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	end if

	'// 모드별 분기
	Select Case mode
		Case "save"
			'등록여부 확인
			sqlStr = "Select count(*) From db_item.dbo.tbl_lotte_brand_mapping Where TenMakerid='" & TenMakerid & "'"
			rsget.Open sqlStr,dbget,1
			if rsget(0)>0 then
				'수정
				sqlStr = "Update db_item.dbo.tbl_lotte_brand_mapping Set " &_
					"	lotteBrandCd='" & lotteBrandCd & "'" &_
					" 	,lotteBrandNm='" & lotteBrandNm & "'" &_
					" Where TenMakerid='" & TenMakerid & "'"
				dbget.execute(sqlStr)
			else
				'신규등록
				sqlStr = "Insert into db_item.dbo.tbl_lotte_brand_mapping values " &_
						" ('" & TenMakerid & "'" &_
						", '" & lotteBrandCd & "','" & lotteBrandNm & "','Y', getdate()) "
				dbget.execute(sqlStr)
			end if
			rsget.Close

		Case "del"
			'매칭된 텐바이텐 카테고리 삭제
			sqlStr = "Delete From db_item.dbo.tbl_lotte_brand_mapping " &_
					" Where TenMakerid='" & TenMakerid & "'"
			dbget.execute(sqlStr)
	End Select
%>
<script language="javascript">
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->