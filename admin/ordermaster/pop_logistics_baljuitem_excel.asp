<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  �¶��� �������
' History : 2020.07.08 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/logistics/logistics_baljuipgocls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljucls.asp" -->
<!-- #include virtual="/lib/BarcodeFunction.asp" -->
<html>
<%
dim baljukey ,section ,obaljupage, sitebaljukey, i, menupos
	baljukey = requestcheckvar(getNumeric(request("baljukey")),10)
	sitebaljukey = requestcheckvar(getNumeric(request("sitebaljukey")),10)
    menupos = requestcheckvar(getNumeric(request("menupos")),10)

set obaljupage = new CBaljuIpgo
	obaljupage.FRectBaljuKey = sitebaljukey
	obaljupage.GetBaljuIpgoitem

Response.Buffer = true    '���ۻ�뿩��
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"

%>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����ڵ�</td>
	<td>����</td>
	<td>�귣��ID</td>
    <td>�����</td>
</tr>
<% if obaljupage.FResultCount>0 then %>
    <% for i = 0 to obaljupage.FResultCount-1 %>
    <tr class="a" height="25" bgcolor="#FFFFFF" align="center">
        <td class='txt'><%= BF_MakeTenBarcode(obaljupage.FItemList(i).fitemgubun,obaljupage.FItemList(i).fitemid,obaljupage.FItemList(i).fitemoption) %></td>
        <td><%= obaljupage.FItemList(i).fitemno %></td>
        <td class='txt'><%= obaljupage.FItemList(i).fmakerid %></td>
        <td class='txt'><%= obaljupage.FItemList(i).FwarehouseCd %></td>
    </tr>
    <%
    if i mod 3000 = 0 then
        Response.Flush		' ���۸��÷���
    end if
    next
    %>
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="4" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>

</table>
</body>
</html>
<%
set obaljupage = Nothing
%>
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
