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

'// ���� ����
Dim oAttrib, itemid, arrDispCate, i
Dim tmpAtrDiv

'// �Ķ���� ����
itemid = requestCheckVar(request("itemid"),10)
arrDispCate = request("arrDispCate")

if itemid="" or arrDispCate="" then
	Response.Write "������ ����ī�װ��� �����ϴ�."
	dbget.Close: Response.End
end if

'// ���������� ���
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
		Response.Write "ī�װ��� ����� ��ǰ�Ӽ��� �����ϴ�."
	end if

	set oAttrib = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->