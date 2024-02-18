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
    dim cardItemId, cardOption
    cardItemId = Request("cardItemId")
    cardOption = Request("cardOption")

	'// 트랜젝션 시작
	dbget.beginTrans

	Select Case mode
		Case "add"		'// 옵션 신규 등록
			'옵션번호 생성
			sqlStr = "Select Max(cardOption) as cardOption From db_item.dbo.tbl_giftcard_option Where cardItemId=" & cardItemId
			rsget.Open sqlStr,dbget,1
				cardOption = Num2Str(cInt(rsget("cardOption"))+1,4,"0","R")
			rsget.close

			sqlStr = "Insert into db_item.dbo.tbl_giftcard_option " & vbCrLf
			sqlStr = sqlStr & "(cardItemId, cardOption, cardOptionName, cardSellCash, cardSalePrice, cardOrgPrice, optSellYn) " & vbCrLf
			sqlStr = sqlStr & "values " & vbCrLf
			sqlStr = sqlStr & "(" & cardItemId & vbCrLf
			sqlStr = sqlStr & ",'" & cardOption & "'" & vbCrLf
			sqlStr = sqlStr & ",'" & html2db(Request("cardOptionName")) & "'" & vbCrLf
			sqlStr = sqlStr & "," & Request("cardSellCash") & ",0," & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & ",'" & Request("optSellYn") & "')" & vbCrLf
			
			dbget.execute(sqlStr)

		Case "modi"	'// 옵션 수정
			sqlStr = "Update db_item.dbo.tbl_giftcard_option Set " & vbCrLf
			sqlStr = sqlStr & " cardOptionName='" & html2db(Request("cardOptionName")) & "'" & vbCrLf
			sqlStr = sqlStr & " ,cardSellCash=" & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & " ,cardOrgPrice=" & Request("cardSellCash") & vbCrLf
			sqlStr = sqlStr & " ,optSellYn='" & Request("optSellYn") & "'" & vbCrLf
			sqlStr = sqlStr & " Where cardItemId=" & cardItemId & vbCrLf
			sqlStr = sqlStr & "		and cardOption='" & cardOption & "'" & vbCrLf

			dbget.execute(sqlStr)

		Case "del"	'// 옵션 삭제
			sqlStr = "Update db_item.dbo.tbl_giftcard_option Set " & vbCrLf
			sqlStr = sqlStr & " optIsUsing='N' " & vbCrLf
			sqlStr = sqlStr & " Where cardItemId=" & cardItemId & vbCrLf
			sqlStr = sqlStr & "		and cardOption='" & cardOption & "'" & vbCrLf

			dbget.execute(sqlStr)

	End Select

	'##### DB 저장 처리 #####
    If Err.Number = 0 Then
    	dbget.CommitTrans				'커밋(정상)
    	Response.Write "<script language='javascript'>" & vbCrLf
    	Response.Write "alert('데이터를 저장하였습니다.');" & vbCrLf
    	Response.Write "top.opener.history.go(0);" & vbCrLf
    	Response.Write "top.self.close();" & vbCrLf
    	Response.Write "</script>"
    	
    Else
        dbget.RollBackTrans				'롤백(에러발생시)
        Call Alert_msg("처리중 에러가 발생했습니다.")
    End If
%>