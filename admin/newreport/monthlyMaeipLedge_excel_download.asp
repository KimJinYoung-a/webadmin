<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*10		' 10��
%>
<%
'###########################################################
' Description : ����ڻ�(����) FIX �����ٿ�ε�
' History : �̻� ����
'			2023.10.11 �ѿ�� ����(csv���� -> �������� �������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
dim yyyymm, placeGubun, PriceGbn, ver, oCMonthlyMaeipLedge, arrList, i
    yyyymm = RequestCheckVar(request("yyyymm"),7)
    placeGubun = RequestCheckVar(request("placeGubun"),1)
    PriceGbn = RequestCheckVar(request("PriceGbn"),1)
    ver = RequestCheckVar(request("ver"),10)

if (ver = "") then
	ver = "V2"
end if

set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge
oCMonthlyMaeipLedge.FCurrPage = 1
oCMonthlyMaeipLedge.FPageSize = 1000000
oCMonthlyMaeipLedge.frectver = ver
oCMonthlyMaeipLedge.frectyyyymm = yyyymm
oCMonthlyMaeipLedge.frectplaceGubun = placeGubun
oCMonthlyMaeipLedge.frectPriceGbn = PriceGbn
oCMonthlyMaeipLedge.GetMaeipLedgeListNotPaging

if oCMonthlyMaeipLedge.FTotalCount>0 then
    arrLIst=oCMonthlyMaeipLedge.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENMonthlyMaeipLedgeList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="50">
		�˻���� : <b><%= oCMonthlyMaeipLedge.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>YYYY-MM</td>
    <td>�����ġ</td>
    <td>�μ�</td>
    <td>��������</td>
    <td>�귣��</td>
    <td>��ǰ����</td>
    <td>��ǰ�ڵ�</td>
    <td>�ɼ��ڵ�</td>
    <td>���ڵ�</td>
    <td>�ܰ�(���)</td>
    <td>���ʼ���</td>
    <td>���ʱݾ�</td>
    <td>�԰����</td>
    <td>�԰�ݾ�</td>
    <td>�̵�����</td>
    <td>�̵��ݾ�</td>
    <td>�Ǹż���</td>
    <td>�Ǹűݾ�</td>
    <td>����������</td>
    <td>�������ݾ�</td>
    <td>��Ÿ������</td>
    <td>��Ÿ���ݾ�</td>
    <td>CS������</td>
    <td>CS���ݾ�</td>
    <td>��������</td>
    <td>�����ݾ�</td>
    <td>�⸻����</td>
    <td>�⸻�ݾ�</td>

	<% if placeGubun <> "S" then %>
		<td>�����԰��</td>
	<% end if %>
	<% if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then %>
		<td>�����԰��(���Ա��к�)</td>
	<% end if %>

    <td>����ڹ�ȣ</td>
    <td>�����</td>
    <td>����ī�װ�</td>
    <td>����ī��1</td>
    <td>����ī��2</td>
    <td>��������</td>
    <td>���͸��Ա���</td>
    <td>��ǰ���Ա���</td>
    <td>�Һ��ڰ�</td>
    <td>�����ǸŰ�</td>
    <td>�����Ǹſ���</td>

	<% if (ver = "DW") then %>
        <td>���ʽ��������밡</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td>��ǰ��������</td>
		<td>�ɼǴ�������</td>
	<% end if %>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td class="txt"><%= arrList(1,i) %></td><% 'YYYY-MM %>
    <td class="txt"><%= trim(arrList(2,i)) %></td><% '�����ġ %>
    <td><%= arrList(3,i) %></td><% '�μ� %>
    <td><%= arrList(4,i) %></td><% '�������� %>
    <td class="txt"><%= arrList(26,i) %></td><% '�귣�� %>
    <td><%= arrList(5,i) %></td><% '��ǰ���� %>
    <td><%= arrList(6,i) %></td><% '��ǰ�ڵ� %>
    <td class="txt"><%= arrList(7,i) %></td><% '�ɼ��ڵ� %>

    <% if (ver = "DW") then %>
        <td class="txt"><%= arrList(44,i) %></td><% '���ڵ� %>
    <% else %>
        <td class="txt"><%= arrList(42,i) %></td><% '���ڵ� %>
    <% end if %>

    <td><%= arrList(28,i) %></td><% '�ܰ�(���) %>
    <td><%= arrList(8,i) %></td><% '���ʼ���(SYS���) %>
    <td><%= arrList(9,i) %></td><% '���ʱݾ�(SYS���) %>
    <td><%= arrList(10,i) %></td><% '�԰���� %>
    <td><%= arrList(11,i) %></td><% '�԰�ݾ� %>
    <td><%= arrList(12,i) %></td><% '�̵����� %>
    <td><%= arrList(13,i) %></td><% '�̵��ݾ� %>
    <td><%= arrList(14,i) %></td><% '�Ǹż��� %>
    <td><%= arrList(15,i) %></td><% '�Ǹűݾ� %>
    <td><%= arrList(16,i) %></td><% '���������� %>
    <td><%= arrList(17,i) %></td><% '�������ݾ� %>
    <td><%= arrList(20,i) %></td><% '��Ÿ������(��:�ν����) %>
    <td><%= arrList(21,i) %></td><% '��Ÿ���ݾ�(��:�ν����) %>
    <td><%= arrList(22,i) %></td><% 'CS������ %>
    <td><%= arrList(23,i) %></td><% 'CS���ݾ� %>
    <td><%= (arrList(8,i) + arrList(10,i)+ arrList(12,i)+ arrList(14,i)+arrList(16,i)+ arrList(18,i)+arrList(20,i) +arrList(22,i)- arrList(24,i))*-1 %></td><% '�������� %>
    <td><%= (arrList(9,i) + arrList(11,i)+ arrList(13,i)+ arrList(15,i)+arrList(17,i)+ arrList(19,i)+arrList(21,i) +arrList(23,i)- arrList(25,i))*-1 %></td><% '�����ݾ� %>
    <td><%= arrList(24,i) %></td><% '�⸻����(�ý������) %>
    <td><%= arrList(25,i) %></td><% '�⸻�ݾ�(�ý������) %>

    <% if placeGubun <> "S" then %>
        <td class="txt"><%= arrList(29,i) %></td><% '�����԰�� %>
    <% end if %>

    <% if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then %>
        <td class="txt"><%= arrList(30,i) %></td><% '�����԰��(���Ա��к�) %>
    <% end if %>

    <td class="txt"><%= arrList(32,i) %></td><% '����ڹ�ȣ %>
    <td><%= arrList(27,i) %></td><% '����� %>
    <td><%= arrList(31,i) %></td><% '����ī�װ� %>
    <td><%= arrList(34,i) %></td><% '����ī��1 %>
    <td><%= arrList(35,i) %></td><% '����ī��2 %>
    <td><%= arrList(36,i) %></td><% '�������� %>
    <td><%= arrList(37,i) %></td><% '���͸��Ա��� %>
    <td><%= arrList(38,i) %></td><% '��ǰ���Ա��� %>
    <td><%= arrList(39,i) %></td><% '�Һ��ڰ� %>
    <td><%= arrList(40,i) %></td><% '�����ǸŰ� %>
    <td><%= arrList(41,i) %></td><% '�����Ǹſ��� %>

    <% if (ver = "DW") then %>
        <td><%= arrList(42,i) %></td><% '��޾�(���ʽ��������밡) %>
        <td><%= arrList(43,i) %></td><% '��ǰ�� %>
        <td><%= arrList(47,i) %></td><% '�ɼǸ� %>
        <td><%= arrList(45,i) %></td><% '��ǰ�������� %>
        <td><%= arrList(46,i) %></td><% '�ɼǴ������� %>
    <% end if %>
</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' ���۸��÷���
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="50" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>

<%
set oCMonthlyMaeipLedge = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->