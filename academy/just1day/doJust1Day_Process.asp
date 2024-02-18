<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doJust1Day_Process.asp
' Discription : 저스트 원데이 처리 페이지
' History : 2016.08.01 유태욱 핑거스
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode, sqlStr, lp
dim justDate, sale_code, itemid, salePrice, orgPrice, saleSuplyCash, limitNo, limitYn, justDesc, img1, img2', , img3, img4
Dim orgsailprice, orgsailsuplycash, orgsailyn

menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)

justDate	= RequestCheckvar(Request("justDate"),10)
sale_code	= getNumeric(Request("sale_code"))
itemid		= getNumeric(Request("itemid"))
salePrice	= getNumeric(Request("salePrice"))
orgPrice	= getNumeric(Request("orgPrice"))
saleSuplyCash = getNumeric(Request("saleSuplyCash"))
limitNo		= getNumeric(Request("limitNo"))
limitYn		= RequestCheckvar(Request("limitYn"),1)
justDesc	= html2db(Request("justDesc"))
img1		= Request("image1")
img2		= Request("image2")
'img3		= Request("image3")
'img4		= Request("image4")
if justDesc <> "" then
	if checkNotValidHTML(justDesc) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if img1 <> "" then
	if checkNotValidHTML(img1) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if img2 <> "" then
	if checkNotValidHTML(img2) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
'// 트랜젝션 시작
dbACADEMYget.beginTrans

'// 모드에 따른 분기
Select Case mode
	Case "add"
		'// 신규 등록
		rsACADEMYget.Open "Select count(JustDate) from [db_academy].[dbo].[tbl_just1day] where JustDate='" & justDate & "'", dbACADEMYget, 1
		if rsACADEMYget(0)>0 then
			Alert_return("이미 등록된 날짜입니다.\n다른 날짜로 변경해주세요.")
			dbACADEMYget.RollBackTrans														'// 2015-04-23, skyer9
			dbACADEMYget.close()	:	response.End
		end if
		rsACADEMYget.Close

		'// 세일된 원가격과 세일여부를 가져옴
		rsACADEMYget.Open " Select top 1 sailprice, sailsuplycash, saleyn From db_academy.dbo.[tbl_diy_item] Where itemid='"&itemid&"' "
		If Not(rsACADEMYget.bof Or rsACADEMYget.eof) Then
			orgsailprice = rsACADEMYget("sailprice")
			orgsailsuplycash = rsACADEMYget("sailsuplycash")
			orgsailyn = rsACADEMYget("saleyn")
		End If
		rsACADEMYget.Close

		'' 할인가 0, 매입가0 인경우 할인 안됨. 일단 등록
		'할인예약 테이블 저장(마스터)
		sqlStr = "Insert Into [db_academy].[dbo].tbl_sale " &_
				" (sale_name, sale_rate, sale_margin, sale_marginvalue, sale_startdate, sale_enddate, availPayType, adminid, sale_status) values " &_
				" ('Just1Day_" & justDate & "' " &_
				" ," & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" , 5, " & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" ,'" & justDate & "', '" & justDate & "'" &_
				" , '8', '" & session("ssBctId") & "',7)"
		dbACADEMYget.Execute(sqlStr)
		sqlStr = "select IDENT_CURRENT('[db_academy].[dbo].tbl_sale') as sale_code"
		rsACADEMYget.Open sqlStr, dbACADEMYget, 1
		If Not rsACADEMYget.Eof then
			sale_code = rsACADEMYget("sale_code")
		end if
		rsACADEMYget.close

		'할인예약 테이블 저장(상품서브)
		sqlStr = "Insert Into [db_academy].[dbo].tbl_saleItem " &_
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
		dbACADEMYget.Execute(sqlStr)

		'저스트 원데이 저장
		sqlStr = "Insert Into [db_academy].[dbo].tbl_just1day " &_
				" (JustDate,itemid,orgPrice,justSalePrice,SaleSuplyCash,justDesc,sale_code,limitNo,adminid, img2, img1) values " &_
				" ('" & justDate & "'" &_
				" ," & itemid &_
				" ," & orgPrice &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ,'" & justDesc & "'" &_
				" ," & sale_code &_
				" ," & limitNo &_
				" ,'" & session("ssBctId") & "'" &_
				" ,'" & img2 & "'" &_
				" ,'" & img1 & "')"

		dbACADEMYget.Execute(sqlStr)

	Case "edit"
		'// 내용 수정
		sqlStr = "Update [db_academy].[dbo].tbl_just1day " &_
				" Set justSalePrice=" & salePrice &_
				" 	,SaleSuplyCash=" & SaleSuplyCash &_
				" 	,limitNo=" & limitNo &_
				" 	,justDesc='" & justDesc & "'" &_
				" 	,img1='" & img1 & "'" &_
				" 	,img2='" & img2 & "'" &_
				" Where justDate='" & justDate & "'"
		dbACADEMYget.Execute(sqlStr)

        if sale_code<>"" then
    		sqlStr = "Update [db_academy].[dbo].tbl_sale " &_
    				" Set sale_rate=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" 	,sale_marginvalue=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" Where sale_code=" & sale_code
    		dbACADEMYget.Execute(sqlStr)


    		sqlStr = "Update [db_academy].[dbo].tbl_saleItem " &_
    				" Set saleprice=" & saleprice &_
    				" 	,salesupplyCash=" & SaleSuplyCash &_
    				" 	,limitno=" & limitno &_
    				"	,lastupdate=getdate() " &_
    				" Where sale_code=" & sale_code & " and itemid=" & itemid
    		dbACADEMYget.Execute(sqlStr)
        end if
	Case "delete"
		'// 삭제
		if justDate>cStr(date()) then
			if sale_code<>"" then
				'관련 할인예약 정보 삭제
				sqlStr = "Update [db_academy].[dbo].tbl_sale " &_
						" Set sale_using=0 " &_
						" Where sale_code=" & sale_code & ";" & vbCrLf
			end if
			'저스트원데이 완전 삭제
			sqlStr = sqlStr & "delete [db_academy].[dbo].tbl_just1day " &_
					" Where justDate='" & justDate & "';" & vbCrLf
			dbACADEMYget.Execute(sqlStr)
		else
			Alert_return("현재 진행중이거나 완료된 상품은 삭제할 수 없습니다.")
			response.End
		end if

End Select


'// 트랜젝션 검사 및 실행
If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
Else
        dbACADEMYget.RollBackTrans
		Alert_return("데이타를 저장하는 도중에 에러가 발생하였습니다.")
		dbACADEMYget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "Just1Day_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
