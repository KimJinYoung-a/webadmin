<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sMode, idx
Dim roomname, MaxSu, orderNo, isusing
Dim strSql

sMode		= Request("sM")
idx			= Request("idx")
roomname 	= Request("roomname")
MaxSu 		= Request("MaxSu")
orderNo 	= Request("orderNo")
isusing 	= Request("isusing")


if (checkNotValidHTML(roomname) = True) Then
	response.write "<script>alert('���̳��Ǹ��� HTML�� ����Ͻ� �� �����ϴ�.');</script>"
	dbget.close()	:	response.End
End If

IF sMode = "I" THEN
	strSql = " INSERT INTO db_partner.dbo.tbl_seminarRoom (roomname, MaxSu, orderNo, isusing)"&_
			" Values('"&roomname&"','"&MaxSu&"','"&orderNo&"','"&isusing&"') "
	dbget.execute strSql

ELSEIF sMode="U" THEN
	strSql =" UPDATE db_partner.dbo.tbl_seminarRoom Set roomname = '"&roomname&"', MaxSu ='"&MaxSu&"', orderNo = '"&orderNo&"', isusing ='"&isusing&"' "&_
			" WHERE idx = '"&idx&"'"
	dbget.execute strSql
END IF

	Response.Write "<script>alert('ó���Ǿ����ϴ�.');location.href='/admin/seminar/popSeminarRoom.asp';</script>"
	dbget.close()
	Response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->