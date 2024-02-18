<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ⱸ�� ����� �߼� ���� �ٿ�ε�
' History : 2016.06.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim itemid, itemoption, i, menupos, page, orderserial, userid, sendstatus, arrlist
dim reserveDlvDate, reserveidx, reserveItemID, reserveItemOption, reserveItemName, regadminid, regdate
dim lastadminid, lastupdate, username, isusing, reloading, jukyogubun
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	reserveitemid = getNumeric(requestcheckvar(request("reserveitemid"),10))
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	itemoption = requestcheckvar(request("itemoption"),4)
	page = getNumeric(requestcheckvar(request("page"),10))
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	orderserial = requestcheckvar(request("orderserial"),11)
	username = requestcheckvar(request("username"),32)
	userid = requestcheckvar(request("userid"),32)
	isusing = requestcheckvar(request("isusing"),1)
	reloading = requestcheckvar(request("reloading"),2)
	sendstatus = requestcheckvar(request("sendstatus"),10)
	jukyogubun = requestcheckvar(request("jukyogubun"),16)

if reloading="" and isusing="" then isusing="Y"
if page="" then page=1

dim ouser
set ouser = new Citemstanding
	ouser.FPageSize = 100000
	ouser.FCurrPage = 1
	ouser.FRectItemID = itemid
	ouser.FRectreserveitemid = reserveitemid
	ouser.FRectitemoption = itemoption
	ouser.FRectreserveidx = reserveidx
	ouser.FRectorderserial = orderserial
	ouser.FRectusername = username
	ouser.FRectuserid = userid
	ouser.FRectisusing = isusing
	ouser.FRectsendstatus = sendstatus
	ouser.FRectjukyogubun = jukyogubun
	ouser.fitemstanding_user_getrows

if ouser.ftotalcount >0 then
	arrlist = ouser.fstandingarr
end if

downPersonalInformation_rowcnt=ouser.ftotalcount

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_���ⱸ��_��۸���Ʈ_" & Left(CStr(now()),10) & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
%>
<!-- #include virtual="/lib/checkAllowIPWithLog_exceldown.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#DDDDDD" border=1>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= ouser.ftotalcount %></b>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
    <td>����ȸ��Vol.(��ȣ)</td>
    <td>��ۻ�ǰ�ڵ�</td>
    <td>��ۿɼ��ڵ�</td>
    <td>��ۻ�ǰ��</td>
	<td>����</td>
    <td>�ֹ���ȣ</td>
    <td>����</td>
    <td>���̵�</td>
    <td>�̸�</td>
	<td>����</td>
	<td>�߼���</td>
	<td>��뿩��</td>
    <td>�Ǹſ��ǰ�ڵ�</td>
    <td>�Ǹſ�ɼ��ڵ�</td>
	<td>�����ȣ</td>
	<td>�ּ�</td>
	<td>���ּ�</td>
	<td>��ȭ��ȣ</td>
	<td>�ڵ���</td>
</tr>

<% if isarray(arrlist) then %>
	<%
	for i=0 to ubound(arrlist,2)
	%>
	<tr bgcolor="<%=chkIIF(arrlist(15,i)="Y","#FFFFFF","#DDDDDD")%>" align="center">
	    <td><%= arrlist(3,i) %></td>
	    <td class='txt'><%= arrlist(21,i) %></td>
		<td class='txt'><%= arrlist(22,i) %></td>
		<td class='txt' align="left"><%= arrlist(23,i) %></td>
		<td><%= getjukyoname(arrlist(4,i)) %></td>
		<td class='txt'><%= arrlist(5,i) %></td>
		<td><%= arrlist(7,i) %></td>
		<td class='txt'><%= arrlist(6,i) %></td>
		<td><%= arrlist(10,i) %></td>
		<td><%= getsendstatusname(arrlist(8,i)) %></td>
		<td>
	    	<%= left(arrlist(9,i),10) %>
	    	<Br><%= mid(arrlist(9,i),12,11) %>
		</td>
		<td><%= arrlist(16,i) %></td>
		<td><%= arrlist(1,i) %></td>
		<td class='txt'><%= arrlist(2,i) %></td>
		<td class='txt' align="left"><%= arrlist(11,i) %></td>
		<td class='txt' align="left"><%= arrlist(12,i) %></td>
		<td class='txt' align="left"><%= arrlist(13,i) %></td>
		<td class='txt'><%= arrlist(14,i) %></td>
		<td class='txt'><%= arrlist(15,i) %></td>
	</tr>
	<%
	if i mod 3000 = 0 then
		Response.Flush		' ���۸��÷���
	end if
	Next
	%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="19" align="center">�˻������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>
</html>

<%
set ouser=nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->