<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���
' History : 2008.05.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim page, i, mailergubun
	page = request("page")
	mailergubun = request("mailergubun")
	
if page="" then page=1

dim omd
set omd = New CMailzine
	omd.FCurrPage = page
	omd.FPageSize=20
	omd.frectmailergubun = mailergubun
	omd.GetMailingList
%>

<link href="/report.css" rel="stylesheet" type="text/css">

<script language="javascript">

//�űԵ�� �˾�����
function popup()
{
	var popup = window.open('/admin/mailopen/mail_reg.asp?mode=add','popup','width=1024,height=768,scrollbars=yes,resizable=yes');
	popup.focus();
}

function popupedit(idx)
{
	var popupedit = window.open('/admin/mailopen/mail_edit.asp?idx='+idx+'&mode=edit','popupedit','width=1024,height=768,scrollbars=yes,resizable=yes');
	popupedit.focus();
	
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �߼۸��Ϸ� : <% drawmailergubun "mailergubun" , mailergubun , "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">	
<tr valign="bottom">       
    <td align="left">
		�� �߼��� 3���� �����Ŀ� �߼���Ȳ�� ��ϵǾ� ���ϴ�.
		<Br>&nbsp;&nbsp;&nbsp;- EMS���Ϸ� ���� ����3��15�п� �ڵ� ���
		<Br>&nbsp;&nbsp;&nbsp;- TMS���Ϸ� ���� ����3�ÿ� �ڵ� ���
    </td>
    <td align="right">
    	<!--<input type="button" value="THUNDERMAIL���" class="button" onclick="javascript:popup();">-->
    </td>        
</tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= omd.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= omd.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�߼� �̸�</td>
	<td>�Ѵ���ڼ�</td>
	<td>�߼۽ð�</td>
	<td>�Ϸ�ð�</td>
	<td>�߼�<br>���Ϸ�</td>
	<td>���</td> 
</tr>
<% if omd.FresultCount>0 then %>
<% for i=0 to omd.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for�� �ȿ��� i ���� ������ ����-->	 
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td><% = omd.FItemList(i).Ftitle %></td>
	<td align="right"><% = FormatNumber(omd.FItemList(i).Ftotalcnt,0) %></td>
	<td><% = omd.FItemList(i).Fstartdate %></td>
	<td><% = omd.FItemList(i).Fenddate %></td>
	<td><% = omd.FItemList(i).fmailergubun %></td>
	<td width=60>
		<input type="button" value="����" onclick="javascript:popupedit(<% = omd.FItemList(i).Fidx %>);" class="button">
	</td>					
</tr>   
</form>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if omd.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= omd.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + omd.StartScrollPage to omd.StartScrollPage + omd.FScrollCount - 1 %>
			<% if (i > omd.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(omd.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if omd.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>	
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	</td>
</tr>
</table>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->