<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%
'###############################################
' Discription : ��ǰ �Ӽ� - ��ǰ ���� ó��
' History : 2019.05.10 ������ : �ű� ����
'###############################################

'// ���� ����
Dim mode, attribCd, i
Dim itemid, itemoption
Dim oJson, sqlStr

'// �Ķ���� ����
mode = requestCheckVar(request("mode"),16)
attribCd = requestCheckVar(request("attribCd"),8)
itemid = requestCheckVar(request("itemid"),8)
itemoption = requestCheckVar(request("itemoption"),4)

'//��� ���
Response.ContentType = "application/json"

'// json��ü ����
Set oJson = jsObject()

if Not(session("ssBctId")<>"") then
	Response.Status = "401 Unauthorized"
	oJson("response") = "Fail"
	oJson("faildesc") = "�߸��� �����Դϴ�."
	oJson.flush
	Set oJson = Nothing
	dbget.close: response.End
end if

if attribCd="" then
	Response.Status = "400 Bad Request"
	oJson("response") = "Fail"
	oJson("faildesc") = "��ǰ�Ӽ������� �����ϴ�."
end if

if itemid="" then
	Response.Status = "400 Bad Request"
	oJson("response") = "Fail"
	oJson("faildesc") = "��ǰ������ �����ϴ�."
end if

on Error Resume Next

Select Case mode
	Case "addLinkItem"
		'// ��ǰ ����
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
		'// ��ǰ ����
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
		'// ���о���
		Response.Status = "400 Bad Request"
		oJson("response") = "Fail"
		oJson("faildesc") = "�߸��� ȣ���Դϴ�."
End Select

IF (Err) then
	Response.Status = "500 Internal Server Error"
	oJson("response") = "Fail"
	oJson("faildesc") = "ó���� ������ �߻��߽��ϴ�."
End if

'Json ���(JSON)
oJson.flush

Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->