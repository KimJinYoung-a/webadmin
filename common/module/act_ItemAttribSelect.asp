<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
Response.CharSet = "euc-kr"

'// 변수 선언
Dim oAttrib, itemid, arrDispCate, i
Dim tmpAtrDiv

'// 파라메터 접수
itemid = requestCheckVar(request("itemid"),10)
arrDispCate = request("arrDispCate")

if itemid="" or arrDispCate="" then
	Response.Write "지정된 전시카테고리가 없습니다."
	dbget.Close: Response.End
end if

'// 페이지정보 목록
	set oAttrib = new CAttrib
	oAttrib.FRectItemid = itemid
	oAttrib.FRectDispCate = arrDispCate
    oAttrib.GetAttribList4Item

	if oAttrib.FResultCount>0 then
		Response.Write "<table cellpadding='2' cellspacing='2' class='a' >"
		Response.Write "<tr>"

		tmpAtrDiv = oAttrib.FItemList(0).FattribDiv
		for i=0 to oAttrib.FResultCount-1
			if i=0 or (tmpAtrDiv<>oAttrib.FItemList(i).FattribDiv) then
				if i>0 then
					Response.Write "</tr><tr>"
				end if

				Response.Write "<td bgcolor='#F0F0F8'>" & oAttrib.FItemList(i).FattribDivName & "</td>"
				tmpAtrDiv = oAttrib.FItemList(i).FattribDiv
				Response.Write "<td bgcolor='#F8F8F8'>"
			end if

			Response.Write "<span style='white-space:nowrap;'><label><input type='checkbox' name='attribCd' value='" & oAttrib.FItemList(i).FattribCd & "' " & chkIIF(oAttrib.FItemList(i).FchkAttrib,"checked","") & "/>" & oAttrib.FItemList(i).FattribName & "</label></span> "

		next
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"
	else
		Response.Write "카테고리에 연결된 상품속성이 없습니다."
	end if

	set oAttrib = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->