<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs����
' History : 2009.04.17 �̻� ����
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->

<%
dim orderserial, makerid
dim i, j, k
	orderserial = requestCheckVar(request("orderserial"),32)
    makerid = requestCheckVar(request("makerid"),32)

dim oAddSongjang
dim IsAddSongjangExist : IsAddSongjangExist = False
set oAddSongjang = new COrderMaster

oAddSongjang.FRectOrderSerial = orderserial
oAddSongjang.GetAddSongjangList()

if (oAddSongjang.FResultCount > 0) then
    IsAddSongjangExist = True
end if

if Not IsAddSongjangExist then
    response.write "�߸��� �����Դϴ�."
    dbget.close() : response.end
end if

%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script type="text/javascript">

function popDeliveryTrace(traceUrl, songjangNo){
	var f = document.popForm;
	f.traceUrl.value	= traceUrl;
	f.songjangNo.value	= songjangNo;
	f.submit();
}

document.title = "�߰����� ����";

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="100">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�߰����� ����</b>
			    </td>
			    <td align="right">

			    </td>
			</tr>
		</table>
    </td>
</tr>
<%
for i = 0 to oAddSongjang.FResultCount - 1
    if ((makerid = "") and (oAddSongjang.FItemList(i).Fmakerid = "")) or _
        ((makerid <> "") and (oAddSongjang.FItemList(i).Fmakerid = makerid)) then
%>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">�����ȣ</td>
    <td>
        <%= oAddSongjang.FItemList(i).Fsongjangdivname %>
        &nbsp;
        <% if (oAddSongjang.FItemList(i).Fsongjangdiv = "24") then %>
        <a href="javascript:popDeliveryTrace('<%= oAddSongjang.FItemList(i).Ffindurl %>','<%= oAddSongjang.FItemList(i).Fsongjangno %>');"><%= oAddSongjang.FItemList(i).Fsongjangno %></a>
        <% else %>
        <a target="_blank" href="<%= oAddSongjang.FItemList(i).Ffindurl + oAddSongjang.FItemList(i).Fsongjangno %>"><%= oAddSongjang.FItemList(i).Fsongjangno %></a>
        <% end if %>
    </td>
</tr>
<%
        ''exit for
    end if
next

%>
</table>

<form name="popForm" action="popDeliveryTrace.asp" target="_blank" style="margin:0px;">
<input type="hidden" name="traceUrl">
<input type="hidden" name="songjangNo">
</form>

<%
set oAddSongjang = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
