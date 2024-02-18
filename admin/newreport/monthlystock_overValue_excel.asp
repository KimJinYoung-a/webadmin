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
' Description : ������ �����ٿ�ε�
' History : �̻� ����
'			 2023.10.11 �ѿ�� ����(csv���� -> �������� �������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
Dim reqYYYYMM, reqStrplace, reqsysorreal, reqbPriceGbn, reqmygubun, reqYYYY, IsUsingV2, strNoType, strPriceType, strYearMonth
Dim AdmPath, appPath, sNow, sY, sM, sD, sH, sMi, sS, sDateName, FileName, fso, tFile, FTotCnt, FTotPage, FCurrPage, sqlStr
dim i, ArrRows, headLine, ojaego, arrLIst, tmpPrice
	reqYYYYMM = RequestCheckVar(request("exYYYY"),4)&"-"&RequestCheckVar(request("exMM"),4)
	reqStrplace = RequestCheckVar(request("stplace"),1)
	reqsysorreal = RequestCheckVar(request("sysorreal"),10)
	reqbPriceGbn = RequestCheckVar(request("bPriceGbn"),1)
	reqmygubun = RequestCheckVar(request("mygubun"),1)
	reqYYYY = RequestCheckVar(request("exYYYY"),4)
	IsUsingV2 = RequestCheckVar(request("v2"),10)

if (IsUsingV2 = "") then
	IsUsingV2 = "Y"
end if

set ojaego = new CMonthlyMaeipLedge
ojaego.FCurrPage = 1
ojaego.FPageSize = 1000000
ojaego.frectreqYYYYMM = reqYYYYMM
ojaego.frectreqStrplace = reqStrplace
ojaego.frectIsUsingV2 = IsUsingV2
ojaego.GetJeagoOverValueListNotPaging

if ojaego.FTotalCount>0 then
    arrLIst=ojaego.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENMonthlyStockList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
	<td colspan="40">
		�˻���� : <b><%= ojaego.FTotalCount %></b>
	</td>
</tr>
<%
strNoType		= "�ǻ�(+�ҷ�)"
strPriceType	= "�ۼ��ø��԰�"
strYearMonth	= "1-3����,4����~6����,7����~12����,1��~2��,2���ʰ�"

if (reqsysorreal = "sys") then
    strNoType = "�ý���"
end if
if (reqbPriceGbn = "V") then
    strPriceType = "��ո��԰�"
end if
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�μ�</td>
    <td>��������</td>
    <td>���Ա���</td>
    <td>�귣��</td>
    <td>����</td>
    <td>����</td>
    <td>��ǰ�ڵ�</td>
    <td>�ɼ��ڵ�</td>
    <td>���ڵ�</td>
    <td>��ǰ��</td>
    <td>�ɼǸ�</td>
    <td>�����԰���</td>
    <td>����(�ý���)</td>
    <td>���ް�(<%= strPriceType %>)</td>

    <% if (reqmygubun = "Y") then %>
        <td><%= reqYYYY %></td>
        <td><%= reqYYYY - 1 %></td>
        <td><%= reqYYYY - 2 %></td>
        <td><%= reqYYYY - 3 %></td>
    <% else %>
        <td>1-3����</td>
        <td>4����~6����</td>
        <td>7����~12����</td>
        <td>1��~2��</td>
        <td>2���ʰ�</td>
    <% end if %>

    <td>NULL</td>
    <td>�հ�</td>
    <td>����ī�װ�</td>
    <td>����ī�װ�</td>
    <td>����ī�װ�</td>
    <td>����ī�װ�</td>
    <td>�Һ��ڰ�</td>
    <td>�����ǸŰ�</td>
    <td>�����Ǹſ���</td>
    <td>��������</td>
    <td>���缾�͸��Ա���</td>
    <td>������Ա���</td>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrList(1,i) %></td>
    <td><%= arrList(2,i) %></td>
    <td><%= arrList(3,i) %></td>
    <td class="txt"><%= arrList(4,i) %></td><% ' �귣�� %>
    <td><%= trim(arrList(12,i)) %></td>
    <td><%= arrList(6,i) %></td>
    <td><%= arrList(7,i) %></td>
    <td class="txt"><%= arrList(8,i) %></td><% ' �ɼ��ڵ� %>
    <td class="txt"><%= arrList(40,i) %></td><% ' ���ڵ� %>
    <td><%= arrList(9,i) %></td>
    <td><%= arrList(10,i) %></td>
    <td class="txt"><%= arrList(11,i) %></td><% ' �����԰��� %>
    <td><%= arrList(13,i) %></td>

    <td>
        <% if (reqbPriceGbn = "V") then %>
            <%= arrList(16,i) %>
            <% tmpPrice = arrList(16,i) %>
        <% else %>
            <%= arrList(15,i) %>
            <% tmpPrice = arrList(15,i) %>
        <% end if %>
    </td>
    <% if (reqsysorreal = "sys") then %>
        <% if (reqmygubun = "Y") then %>
            <td><%= arrList(22,i)*tmpPrice %></td>
            <td><%= arrList(23,i)*tmpPrice %></td>
            <td><%= arrList(24,i)*tmpPrice %></td>
            <td><%= arrList(25,i)*tmpPrice %></td>
        <% else %>
            <td><%= arrList(17,i)*tmpPrice %></td>
            <td><%= arrList(18,i)*tmpPrice %></td>
            <td><%= arrList(19,i)*tmpPrice %></td>
            <td><%= arrList(20,i)*tmpPrice %></td>
            <td><%= arrList(21,i)*tmpPrice %></td>
        <% end if %>

        <td><%= arrList(26,i)*tmpPrice %></td>
        <td><%= arrList(13,i)*tmpPrice %></td>
        <td><%= arrList(38,i) %></td>
    <% else %>
        <% if (reqmygubun = "Y") then %>
            <td><%= arrList(22+10,i)*tmpPrice %></td>
            <td><%= arrList(23+10,i)*tmpPrice %></td>
            <td><%= arrList(24+10,i)*tmpPrice %></td>
            <td><%= arrList(25+10,i)*tmpPrice %></td>
        <% else %>
            <td><%= arrList(17+10,i)*tmpPrice %></td>
            <td><%= arrList(18+10,i)*tmpPrice %></td>
            <td><%= arrList(19+10,i)*tmpPrice %></td>
            <td><%= arrList(20+10,i)*tmpPrice %></td>
            <td><%= arrList(21+10,i)*tmpPrice %></td>
        <% end if %>

        <td><%= arrList(26+10,i)*tmpPrice %></td>
        <td><%= arrList(13+1,i)*tmpPrice %></td>
        <td><%= arrList(38,i) %></td>
    <% end if %>
    
    <td><%= arrList(41,i) %></td>

    <% ' ����ī�װ� %>
    <td><%= arrList(42,i) %></td>
    <td><%= arrList(43,i) %></td>
    <td><%= arrList(44,i) %></td>
    <td><%= arrList(45,i) %></td>
    <td><%= arrList(46,i) %></td>
    <td><%= arrList(47,i) %></td>
    <td><%= arrList(48,i) %></td>
    <td><%= arrList(49,i) %></td>
</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' ���۸��÷���
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="40" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->