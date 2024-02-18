<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

Server.ScriptTimeOut = 60*10		' 10��
%>
<%
'###########################################################
' Description : ���� ȭ���� �޸� ��ȯ ������
' History : 2023.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/crm/DormantSleepCls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1, fromDate,i
    page = RequestCheckVar(getNumeric(request("page")),10)
    research = RequestCheckVar(request("research"),2)
    yyyy1 = RequestCheckVar(request("yyyy1"),4)
    mm1   = RequestCheckVar(request("mm1"),2)
    dd1   = RequestCheckVar(request("dd1"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-1,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-1,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-1,date())))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))

dim odormantsleep
set odormantsleep = new CDormantSleepList
    odormantsleep.FCurrPage = page
    odormantsleep.FPageSize = 100
    odormantsleep.FRectStartDate = fromDate
    odormantsleep.GetDormantSleepList
%>
<script type='text/javascript'>

function NextPage(page){
	document.frm.target = "";
	document.frm.action = "";
    document.frm.page.value=page;
    document.frm.submit();
}

function DormantSleepdownloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/crm/DormantSleep_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
		* �޸���ȯ������¥ :
        <% DrawOneDateBoxdynamic "yyyy1", yyyy1, "mm1", mm1, "dd1", dd1, "", "yyyy1", "mm1", "dd1" %>
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="NextPage('1');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left"></td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        �� <%= CStr(DateSerial(yyyy1, mm1, dd1)) %>�Ͽ� �޸����� ��ȯ�� ������ �� ����Ʈ �Դϴ�.
    </td>
    <td align="right">
		<input type="button" onclick="DormantSleepdownloadexcel();" value="�����ٿ�ε�" class="button">
    </td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%= odormantsleep.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= odormantsleep.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�����̵�</td>
    <td>����</td>
    <td>ȸ�����</td>
    <td>Ǫ�ü���</td>
    <td>���ڼ���</td>
    <td>�̸��ϼ���</td>
    <td>�������α���</td>
    <td>���縶�ϸ���</td>
    <td>���</td>
</tr>
<% if odormantsleep.FResultCount>0 then %>
<% for i=0 to odormantsleep.FResultCount-1 %>
<tr bgcolor="#FFFFFF" align="center">
    <td>
        <% if C_CriticInfoUserLV1 then %>
            <%= odormantsleep.FItemList(i).fuserid %>
        <% else %>
            <%= printUserId(odormantsleep.FItemList(i).fuserid,2,"*") %>
        <% end if %>
    </td>
    <td>
        <% if C_CriticInfoUserLV1 then %>
            <%= odormantsleep.FItemList(i).fusername %>
        <% else %>
            <%= printUserId(odormantsleep.FItemList(i).fusername,2,"*") %>
        <% end if %>
    </td>
    <td><%= odormantsleep.FItemList(i).fuserlevel %></td>
    <td><%= odormantsleep.FItemList(i).fpushYn %></td>
    <td><%= odormantsleep.FItemList(i).fsmsok %></td>
    <td><%= odormantsleep.FItemList(i).femailok %></td>
    <td><%= odormantsleep.FItemList(i).flastlogin %></td>
    <td><%= FormatNumber(odormantsleep.FItemList(i).fcurrentMileage, 0) %></td>
    <td></td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="9" align="center">
		<% if odormantsleep.HasPreScroll then %>
		<a href="javascript:NextPage('<%= odormantsleep.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + odormantsleep.StartScrollPage to odormantsleep.FScrollCount + odormantsleep.StartScrollPage - 1 %>
			<% if i>odormantsleep.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if odormantsleep.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set odormantsleep = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
