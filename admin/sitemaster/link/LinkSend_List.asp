<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2019.10.16 �ѿ�� ����
'	Description : Link �߼�
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/rndSerial.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/LinkSendCls.asp"-->
<%
dim page, linkidx, title, isusing, i
	page = requestCheckVar(getNumeric(request("page")),10)
	linkidx = requestCheckVar(getNumeric(request("linkidx")),10)
	title = requestCheckVar(request("title"),128)
	isusing = requestCheckVar(request("isusing"),1)

if page="" then page=1
if isusing="" then isusing="Y"

dim oLink
set oLink = New CLinkSend
    oLink.FCurrPage = page
    oLink.FPageSize=20
    oLink.FRectlinkidx = linkidx
    oLink.FRecttitle = title
    oLink.FRectisusing = isusing
    oLink.GetLinkSend

%>
<script type='text/javascript'>

function popsendlist(linkidx){
	var popwin = window.open('/admin/sitemaster/link/LinkSend_reg.asp?linkidx='+linkidx,'addreg','width=1400,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function gotoPage(pg) {
    document.Listfrm.page.value=pg;
    document.Listfrm.submit();
}

function fnLinkURLCopy(link) {
	window.clipboardData.setData("Text", link);
	alert('��ũ�� ����Ǿ����ϴ�.\n���Ͻô� ���� Ctrl+V �Ͻø�˴ϴ�.');
}

</script>

<!-- �˻��� ���� -->
<form name="Listfrm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		* ��ũ��ȣ :
		<input type="text" name="linkidx" size="10" value="<%= linkidx %>">
        &nbsp;
		* ��ũ�� :
		<input type="text" name="title" size="25" value="<%= title %>">
		&nbsp;
        * ��뿩�� : <% drawSelectBoxisusingYN "isusing",isusing,"" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" onclick="gotoPage('1');" value="�˻�">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right">
        <input type="button" value="�űԵ��" onclick="popsendlist('')" class="button">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="18">
        �˻���� : <b><%= oLink.FTotalCount %></b>
        &nbsp;
        ������ : <b><%= page %>/ <%= oLink.FTotalPage %></b>
    </td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=60>��ũ��ȣ</td>
	<td width=220>��ũ��</td>
	<td>�ܺγ��⸵ũ</td>
	<td>������ũ</td>
	<td width=50>��뿩��</td>
	<td width=60>Ŭ����</td>
	<td width=80>��������</td>
    <td width=40>���</td>
</tr>
<% if oLink.FResultCount>0 then %>
<% for i=0 to oLink.FResultCount-1 %>
<tr bgcolor="<%=chkIIF(oLink.FItemList(i).fisusing="Y","#FFFFFF","#E0E0E0")%>">
	<td align="center"><%= oLink.FItemList(i).flinkidx %></td>
	<td align="left"><%= chrbyte(ReplaceBracket(oLink.FItemList(i).ftitle),30,"Y") %></td>
	<td align="left">
		http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(oLink.FItemList(i).flinkidx) %>
		<input type="button" value="��ũ����" onclick="fnLinkURLCopy('http://www.10x10.co.kr/apps/Link/LinkSend.asp?key=<%= rdmSerialEnc(oLink.FItemList(i).flinkidx) %>')" class="button">
	</td>
	<td align="left"><%= ReplaceBracket(oLink.FItemList(i).flinkurl) %></td>
	<td align="center"><%= oLink.FItemList(i).fisusing %></td>
	<td align="center"><%= oLink.FItemList(i).fviewcount %></td>
	<td align="center">
        <%= left(oLink.FItemList(i).flastupdate,10) %>
        <br>
        <%= mid(oLink.FItemList(i).flastupdate,11,22) %>
        <% if oLink.FItemList(i).flastadminid <> "" then %>
            <br>(<%= oLink.FItemList(i).flastadminid %>)
        <% end if %>
    </td>
	<td align="center">
        <input type="button" value="����" onclick="popsendlist('<%= oLink.FItemList(i).flinkidx %>')" class="button">
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">
	<% if oLink.HasPreScroll then %>
		<a href="javascript:gotoPage(<%= oLink.StarScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oLink.StarScrollPage to oLink.FScrollCount + oLink.StarScrollPage - 1 %>
		<% if i>oLink.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:gotoPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oLink.HasNextScroll then %>
		<a href="javascript:gotoPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="18" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
<% set oLink = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->