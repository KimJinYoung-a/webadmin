<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doJust1Day_Process.asp
' Discription : 신한은행용 저스트 원데이 처리 페이지
' History : 2009.10.27 허진원 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, sqlStr, lp
dim justDate, itemid, salePrice, orgPrice, saleSuplyCash, limitNo, justDesc

menupos		= Request("menupos")
mode		= Request("mode")

justDate	= Request("justDate")
itemid		= Request("itemid")
salePrice	= Request("salePrice")
orgPrice	= Request("orgPrice")
saleSuplyCash = Request("saleSuplyCash")
limitNo		= Request("limitNo")
justDesc	= html2db(Request("justDesc"))

'// 트랜젝션 시작
dbget.beginTrans

'// 모드에 따른 분기
Select Case mode
	Case "add"
		'// 신규 등록
		rsget.Open "Select count(JustDate) from db_temp.[dbo].tbl_just1Day_Shinhan where JustDate='" & justDate & "'", dbget, 1
		if rsget(0)>0 then
			Alert_return("이미 등록된 날짜입니다.\n다른 날짜로 변경해주세요.")
			dbget.close()	:	response.End
		end if
		rsget.Close

		'저스트 원데이 저장
		sqlStr = "Insert Into db_temp.[dbo].tbl_just1Day_Shinhan " &_
				" (JustDate,itemid,orgPrice,justSalePrice,SaleSuplyCash,justDesc,limitNo,adminid) values " &_
				" ('" & justDate & "'" &_
				" ," & itemid &_
				" ," & orgPrice &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ,'" & justDesc & "'" &_
				" ," & limitNo &_
				" ,'" & session("ssBctId") & "')"
		dbget.Execute(sqlStr)

	Case "edit"
		'// 내용 수정
		sqlStr = "Update db_temp.[dbo].tbl_just1Day_Shinhan " &_
				" Set justSalePrice=" & salePrice &_
				" 	,SaleSuplyCash=" & SaleSuplyCash &_
				" 	,limitNo=" & limitNo &_
				" 	,justDesc='" & justDesc & "'" &_
				" Where justDate='" & justDate & "'"
		dbget.Execute(sqlStr)
        
	Case "delete"
		'// 삭제
		if justDate>cStr(date()) then
			'저스트원데이 완전 삭제
			sqlStr = sqlStr & "delete db_temp.[dbo].tbl_just1Day_Shinhan " &_
					" Where justDate='" & justDate & "';" & vbCrLf
			dbget.Execute(sqlStr)
		else
			Alert_return("현재 진행중이거나 완료된 상품은 삭제할 수 없습니다.")
			dbget.close()	:	response.End
		end if

End Select


'// 트랜젝션 검사 및 실행
If Err.Number = 0 Then
        dbget.CommitTrans
Else
        dbget.RollBackTrans
		Alert_return("데이타를 저장하는 도중에 에러가 발생하였습니다.")
		dbget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "Just1Day_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->