<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
'###############################################
' Discription : 상품 속성 - 상품 목록 Ajax
' History : 2019.04.29 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim mode, attribCd, includeOption, page, i
Dim dispCate, itemid, makerid, itemname
Dim oAttrib, arrItems
Dim oJson

'// 파라메터 접수
mode = requestCheckVar(request("mode"),12)
attribCd = requestCheckVar(request("attribCd"),8)
includeOption = requestCheckVar(request("includeOption"),1)
page = requestCheckVar(request("page"),8)
dispCate = requestCheckVar(request("disp"),18)
itemid = requestCheckVar(request("itemid"),255)
makerid = requestCheckVar(request("makerid"),32)
itemname = requestCheckVar(request("itemname"),32)

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

if page="" then page=1

'//헤더 출력
Response.ContentType = "application/json"

'// json객체 선언
Set oJson = jsObject()

if Not(session("ssBctId")<>"") then
	Response.Status = "401 Unauthorized"
	oJson("response") = "Fail"
	oJson("faildesc") = "잘못된 접속입니다."
	oJson.flush
	Set oJson = Nothing
	dbget.close: response.End
end if

on Error Resume Next

Select Case mode
	Case "findItem"
		'// 연결 안된 상품 목록 접수
		if attribCd<>"" then
			set oAttrib = new CAttrib
				oAttrib.FRectattribCd = attribCd
				oAttrib.FRectIncludeOption = includeOption
				oAttrib.FRectDispCate = dispCate
				oAttrib.FRectItemid = itemid
				oAttrib.FRectItemName = itemname
				oAttrib.FRectMakerid = makerid
				oAttrib.FPageSize = 50
				oAttrib.FCurrPage = page
				oAttrib.GetNotLinkedItemList

				oJson("response") = "Ok"
				oJson("totalCount") = oAttrib.FTotalCount
				oJson("totalPage") = oAttrib.FTotalpage
				oJson("resultCount") = oAttrib.FResultCount
				Set oJson("items")= jsArray()

				for i=0 to oAttrib.FResultCount-1
					set arrItems = jsObject()

					arrItems("itemid") = oAttrib.FItemList(i).Fitemid
					arrItems("itemname") = oAttrib.FItemList(i).Fitemname
					arrItems("itemoption") = oAttrib.FItemList(i).Fitemoption
					arrItems("optionname") = oAttrib.FItemList(i).Foptionname

					set oJson("items")(null) = arrItems
					set arrItems = Nothing
				next

			set oAttrib = Nothing
		else
			Response.Status = "400 Bad Request"
			oJson("response") = "Fail"
			oJson("faildesc") = "상품속성정보가 없습니다."
		end if

	Case "linkedItem"
		'// 속성 연결 상품 목록 접수
		if attribCd<>"" then
			set oAttrib = new CAttrib
				oAttrib.FRectattribCd = attribCd
				oAttrib.FRectIncludeOption = includeOption
				oAttrib.FRectDispCate = dispCate
				oAttrib.FRectItemid = itemid
				oAttrib.FRectItemName = itemname
				oAttrib.FRectMakerid = makerid
				oAttrib.FPageSize = 50
				oAttrib.FCurrPage = page
				oAttrib.GetLinkedItemList

				oJson("response") = "Ok"
				oJson("totalCount") = oAttrib.FTotalCount
				oJson("totalPage") = oAttrib.FTotalpage
				oJson("resultCount") = oAttrib.FResultCount
				Set oJson("items")= jsArray()

				for i=0 to oAttrib.FResultCount-1
					set arrItems = jsObject()

					arrItems("itemid") = oAttrib.FItemList(i).Fitemid
					arrItems("itemname") = oAttrib.FItemList(i).Fitemname
					arrItems("itemoption") = oAttrib.FItemList(i).Fitemoption
					arrItems("optionname") = oAttrib.FItemList(i).Foptionname

					set oJson("items")(null) = arrItems
					set arrItems = Nothing
				next

			set oAttrib = Nothing
		else
			Response.Status = "400 Bad Request"
			oJson("response") = "Fail"
			oJson("faildesc") = "상품속성정보가 없습니다."
		end if
	Case else
		'// 구분없음
		Response.Status = "400 Bad Request"
		oJson("response") = "Fail"
		oJson("faildesc") = "잘못된 호출입니다."
End Select

IF (Err) then
	Response.Status = "500 Internal Server Error"
	oJson("response") = "Fail"
	oJson("faildesc") = "처리중 오류가 발생했습니다."
End if

'Json 출력(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->