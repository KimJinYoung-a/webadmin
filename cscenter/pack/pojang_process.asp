<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'#######################################################
'	History	:  2015.11.09 �ѿ�� ����
'	Description : ���� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/pack_cls.asp"-->

<%
dim mode ,title, message, sqlStr, i, orderserial, midx
	mode = requestCheckVar(request("mode"),16)
    title = request("title")
    message = request("message")
	orderserial = requestcheckvar(request("orderserial"),11)
	midx = requestCheckVar(request("midx"),10)

dim refip
	refip = request.ServerVariables("HTTP_REFERER")

if (InStr(refip,"10x10.co.kr")<1) then
	response.write "<script type='text/javascript'>alert('�������� ���� ��ΰ� �ƴմϴ�.');</script>"
	dbget.close()	:	response.end
end if

'//�������� ����
if mode="editpojang" then
	if midx="" then
		response.write "<script type='text/javascript'>alert('�ϷĹ�ȣ�� �����ϴ�.'); location.replace('"& refip &"');</script>"
		dbget.close()	:	response.end
	end if
	midx = trim(midx)

	if title<>"" then
		if checkNotValidHTML(title) then
			response.write "<script type='text/javascript'>alert('��������� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���.'); location.replace('"& refip &"');</script>"
			dbget.close()	:	response.end
		end if
	end if
	if message<>"" then
		if checkNotValidHTML(message) then
			response.write "<script type='text/javascript'>alert('�����޼����� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���.'); location.replace('"& refip &"');</script>"
			dbget.close()	:	response.end
		end if
	end if

	'//������ ���̺� ����
    sqlStr = "update db_order.dbo.tbl_order_pack_master" + vbcrlf
    sqlStr = sqlStr & " set title='"& html2db(title) &"'" + vbcrlf
    sqlStr = sqlStr & " , message='"& html2db(message) &"' where" + vbcrlf
    sqlStr = sqlStr & " midx="& midx &""

	'response.write sqlStr & "<br>"
    dbget.Execute sqlStr

	response.write "<script type='text/javascript'>"
	response.write "	alert('���� �Ϸ� �Ǿ����ϴ�.');"
	response.write "	location.replace('/cscenter/pack/pojang_view.asp?orderserial="& orderserial &"&midx="& midx &"');"
	response.write "</script>"
	dbget.close()	:	response.end

else
	'response.write "<script type='text/javascript'>location.replace('"& SSLURL &"/inipay/pack/pack_step1.asp');</script>"
	response.write "<script type='text/javascript'>alert('�������� ��ΰ� �ƴմϴ�.');</script>"
	dbget.close()	:	response.end
end if

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->