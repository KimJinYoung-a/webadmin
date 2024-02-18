<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �˸�����
' Hieditor : 2023.03.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->

<%
Dim page,  searchRect, searchStr, isusing, i, research, serverDownYN, serverEnginDownYN, dbBusyYN, tenbytenWMASiteErrorYN
dim statediv
	page			= requestCheckvar(getNumeric(Request("page")),10)
	searchRect		= requestCheckvar(Request("searchRect"),32)
	searchStr		= requestCheckvar(Request("searchStr"),32)
	research		= requestCheckvar(Request("research"),2)
    statediv			= requestCheckvar(Request("statediv"),1)

if page="" then page=1
isusing = "Y"
if research="" and statediv="" then
	statediv = "Y"
end if

dim cUser
Set cUser = new CUserNotification
    cUser.FPagesize = 20
    cUser.FCurrPage = page
    cUser.FRectSearchRect = searchRect
    cUser.FRectSearchStr = searchStr
    cUser.fRectIsusing = isusing
    cUser.fRectstatediv = statediv
    cUser.GetUserList()
%>

<script type="text/javascript">

function jsGoPage(pg){
	document.frmUser.page.value=pg;
	document.frmUser.submit();
}

function NotificationUser(userId){
	var NotificationUserPop = window.open("/admin/member/notification/NotificationUser.asp?userId=" + userId + "&menupos=<%=menupos%>","NotificationUser","width=1400,height=800,scrollbars=yes");
	NotificationUserPop.focus();
}

</script>

<form name="frmUser" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        * �������� : <% drawSelectBoxisusingYN "statediv", statediv, "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="jsGoPage('');">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
		* �˻����� :
        <select class="select" name="searchRect">
			<option value="">��ü</option>
			<option value="userid" <%= CHKIIF(searchRect="userid", "selected", "") %> >�������̵�</option>
		</select>
		<input type="text" class="text" name="searchStr" value="<%= searchStr %>" size="20">
    </td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
    </td>
    <td align="right">	
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="14">
        �˻���� : <b><%= cUser.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %>/ <%= cUser.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�������̵�</td>
    <td>������</td>
    <td>�����ȣ</td>
    <td>�˸���</td>
    <td>��������</td>
    <td>���</td>
</tr>
<% if cUser.FresultCount>0 then %>
    <% for i=0 to cUser.FresultCount-1 %>
        <% if cUser.FItemList(i).fstatediv = "N" or cUser.FItemList(i).fisusing = 0 then %>
            <tr align="center" bgcolor="#EEEEEE">
        <% else %>    
            <tr align="center" bgcolor="#FFFFFF">
        <% end if %>
        <td><%= cUser.FItemList(i).fuserid %></td>
        <td><%= cUser.FItemList(i).fusername %></td>
        <td><%= cUser.FItemList(i).fempno %></td>
        <td><%= cUser.FItemList(i).fuserCount %></td>
        <td><%= cUser.FItemList(i).fstatediv %></td>
        <td>
            <input type="button" class="button" value="����" onClick="NotificationUser('<%= cUser.FItemList(i).fuserid %>');">
        </td>
    </tr>
    <% next %>

    <tr height="25" bgcolor="FFFFFF">
        <td colspan="14" align="center">
            <% if cUser.HasPreScroll then %>
                <span class="list_link"><a href="#" onclick="jsGoPage(<%= cUser.StartScrollPage-1 %>); return false;">[pre]</a></span>
            <% else %>
            [pre]
            <% end if %>
            <% for i = 0 + cUser.StartScrollPage to cUser.StartScrollPage + cUser.FScrollCount - 1 %>
                <% if (i > cUser.FTotalpage) then Exit for %>
                <% if CStr(i) = CStr(cUser.FCurrPage) then %>
                <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                <% else %>
                <a href="#" onclick="jsGoPage(<%= i %>); return false;" class="list_link"><font color="#000000"><%= i %></font></a>
                <% end if %>
            <% next %>
            <% if cUser.HasNextScroll then %>
                <span class="list_link"><a href="#" onclick="jsGoPage(<%= i %>); return false;">[next]</a></span>
            <% else %>
            [next]
            <% end if %>
        </td>
    </tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="14" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>

<%
set cUser=Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
