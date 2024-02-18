<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doJust1Day_Process.asp
' Discription : 저스트 원데이 처리 페이지
' History : 2008.04.09 허진원 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, sqlStr, lp
dim justDate, sale_code, itemid, salePrice, orgPrice, saleSuplyCash, limitNo, limitYn, justDesc, img1, img2, img3, img4
Dim orgsailprice, orgsailsuplycash, orgsailyn, orgitemdiv

menupos		= Request("menupos")
mode		= Request("mode")

justDate	= Request("justDate")
sale_code	= getNumeric(Request("sale_code"))
itemid		= getNumeric(Request("itemid"))
salePrice	= getNumeric(Request("salePrice"))
orgPrice	= getNumeric(Request("orgPrice"))
saleSuplyCash = getNumeric(Request("saleSuplyCash"))
limitNo		= getNumeric(Request("limitNo"))
limitYn		= Request("limitYn")
justDesc	= html2db(Request("justDesc"))
img1		= Request("image1")
img2		= Request("image2")
img3		= Request("image3")
img4		= Request("image4")

'// 트랜젝션 시작
dbget.beginTrans

'// 모드에 따른 분기
Select Case mode
	Case "add"
		'// 신규 등록
		rsget.Open "Select count(JustDate) from [db_sitemaster].[dbo].tbl_just1day_temp where JustDate='" & justDate & "'", dbget, 1
		if rsget(0)>0 then
			Alert_return("이미 등록된 날짜입니다.\n다른 날짜로 변경해주세요.")
			dbget.RollBackTrans														'// 2015-04-23, skyer9
			dbget.close()	:	response.End
		end if
		rsget.Close

		'// 세일된 원가격과 세일여부를 가져옴
		rsget.Open " Select top 1 sailprice, sailsuplycash, sailyn, itemdiv From db_item.dbo.tbl_item Where itemid='"&itemid&"' "
		If Not(rsget.bof Or rsget.eof) Then
			orgsailprice = rsget("sailprice")
			orgsailsuplycash = rsget("sailsuplycash")
			orgsailyn = rsget("sailyn")
			orgitemdiv = rsget("itemdiv")
		End If
		rsget.Close

	If orgitemdiv<>"21" Then '딜 상품이 아닐때 세일 정보 등록
		'' 할인가 0, 매입가0 인경우 할인 안됨. 일단 등록
		'할인예약 테이블 저장(마스터)
		sqlStr = "Insert Into [db_event].[dbo].tbl_sale " &_
				" (sale_name, sale_rate, sale_margin, sale_marginvalue, sale_startdate, sale_enddate, availPayType, adminid, sale_status) values " &_
				" ('Just1Day_" & justDate & "' " &_
				" ," & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" , 5, " & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" ,'" & justDate & "', '" & justDate & "'" &_
				" , '8', '" & session("ssBctId") & "',7)"
		dbget.Execute(sqlStr)
		sqlStr = "select IDENT_CURRENT('[db_event].[dbo].tbl_sale') as sale_code"
		rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof then
			sale_code = rsget("sale_code")
		end if
		rsget.close

		'할인예약 테이블 저장(상품서브)
		sqlStr = "Insert Into [db_event].[dbo].tbl_saleItem " &_
				" (sale_code, itemid, saleprice, salesupplyCash, limitno, orgsailprice, orgsailsuplycash, orgsailyn,  orglimityn, saleItem_status) values " &_
				" (" & sale_code &_
				" ," & itemid &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ," & limitNo &_
				" ," & orgsailprice &_
				" ," & orgsailsuplycash &_
				" ,'" & orgsailyn &_
				"' ,'" & limitYn & "', 7)"
		dbget.Execute(sqlStr)
	Else
		sale_code=0
	End If
		'저스트 원데이 저장
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_just1day_temp " &_
				" (JustDate,itemid,orgPrice,justSalePrice,SaleSuplyCash,justDesc,sale_code,limitNo,adminid,img1,img2,img3,img4) values " &_
				" ('" & justDate & "'" &_
				" ," & itemid &_
				" ," & orgPrice &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ,'" & justDesc & "'" &_
				" ," & sale_code &_
				" ," & limitNo &_
				" ,'" & session("ssBctId") & "'" &_
				" ,'" & img1 & "','" & img2 & "','" & img3 & "','" & img4 & "')"

		dbget.Execute(sqlStr)

	Case "edit"
		'// 내용 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_just1day_temp " &_
				" Set justSalePrice=" & salePrice &_
				" 	,SaleSuplyCash=" & SaleSuplyCash &_
				" 	,limitNo=" & limitNo &_
				" 	,justDesc='" & justDesc & "'" &_
				" 	,img1='" & img1 & "'" &_
				" 	,img2='" & img2 & "'" &_
				" 	,img3='" & img3 & "'" &_
				" 	,img4='" & img4 & "'" &_
				" Where justDate='" & justDate & "'"
		dbget.Execute(sqlStr)

        If sale_code > 0 Then
    		sqlStr = "Update [db_event].[dbo].tbl_sale " &_
    				" Set sale_rate=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" 	,sale_marginvalue=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" Where sale_code=" & sale_code
    		dbget.Execute(sqlStr)

    		sqlStr = "Update [db_event].[dbo].tbl_saleItem " &_
    				" Set saleprice=" & saleprice &_
    				" 	,salesupplyCash=" & SaleSuplyCash &_
    				" 	,limitno=" & limitno &_
    				"	,lastupdate=getdate() " &_
    				" Where sale_code=" & sale_code & " and itemid=" & itemid
    		dbget.Execute(sqlStr)
        end if
	Case "delete"
		'// 삭제
		if justDate>cStr(date()) then
			if sale_code<>"" then
				'관련 할인예약 정보 삭제
				sqlStr = "Update [db_event].[dbo].tbl_sale " &_
						" Set sale_using=0 " &_
						" Where sale_code=" & sale_code & ";" & vbCrLf
			end if
			'저스트원데이 완전 삭제
			sqlStr = sqlStr & "delete [db_sitemaster].[dbo].tbl_just1day_temp " &_
					" Where justDate='" & justDate & "';" & vbCrLf
			dbget.Execute(sqlStr)
		else
			Alert_return("현재 진행중이거나 완료된 상품은 삭제할 수 없습니다.")
			response.End
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
