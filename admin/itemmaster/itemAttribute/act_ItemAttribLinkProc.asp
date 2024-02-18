<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
'###############################################
' Discription : 상품 속성 - 상품 연결 처리
' History : 2019.05.10 허진원 : 신규 생성
'###############################################

'// 변수 선언
Dim mode, attribCd, i
Dim itemid, itemoption
Dim oJson, sqlStr

'// 파라메터 접수
mode = requestCheckVar(request("mode"),16)
attribCd = requestCheckVar(request("attribCd"),8)
itemid = requestCheckVar(request("itemid"),8)
itemoption = requestCheckVar(request("itemoption"),4)

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

if attribCd="" then
	Response.Status = "400 Bad Request"
	oJson("response") = "Fail"
	oJson("faildesc") = "상품속성정보가 없습니다."
end if

if itemid="" then
	Response.Status = "400 Bad Request"
	oJson("response") = "Fail"
	oJson("faildesc") = "상품정보가 없습니다."
end if

on Error Resume Next

Select Case mode
	Case "addLinkItem"
		'// 상품 연결
		sqlStr = "IF NOT EXISTS( "
        sqlStr = sqlStr & " Select attribCd "
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_itemAttrib_item "
        sqlStr = sqlStr & " 	where attribCd=" & attribCd
        sqlStr = sqlStr & " 		and itemid=" & itemid
		sqlStr = sqlStr & " 		and isNull(itemoption,'')='" & itemoption & "' )"
        sqlStr = sqlStr & " BEGIN "
        sqlStr = sqlStr & " 	insert into db_item.dbo.tbl_itemAttrib_item values "
        sqlStr = sqlStr & " 	("& attribCd &","& itemid &",'"& itemoption &"') "
        sqlStr = sqlStr & " END "
		dbget.execute(sqlStr)

		oJson("response") = "Ok"

	Case "clearLinkItem"
		'// 상품 해제
		sqlStr = "IF EXISTS( "
        sqlStr = sqlStr & " Select attribCd "
        sqlStr = sqlStr & " 	from db_item.dbo.tbl_itemAttrib_item "
        sqlStr = sqlStr & " 	where attribCd=" & attribCd
        sqlStr = sqlStr & " 		and itemid=" & itemid
		sqlStr = sqlStr & " 		and isNull(itemoption,'')='" & itemoption & "' )"
        sqlStr = sqlStr & " BEGIN "
        sqlStr = sqlStr & " 	Delete from db_item.dbo.tbl_itemAttrib_item "
        sqlStr = sqlStr & " 	where attribCd=" & attribCd
        sqlStr = sqlStr & " 		and itemid=" & itemid
		sqlStr = sqlStr & " 		and isNull(itemoption,'')='" & itemoption & "' "
        sqlStr = sqlStr & " END "
		dbget.execute(sqlStr)

		oJson("response") = "Ok"
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