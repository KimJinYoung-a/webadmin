<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������
' Hieditor : 2009.11.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_coincls.asp"-->

<%
dim research,userid, fixtype, linktype, poscode, validdate
dim page

	userid = request("userid")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

if page = "" then page = 1

	dim cMomoBonusCoinList, PageSize , ttpgsz , CurrPage, i
	CurrPage = requestCheckVar(request("cpg"),9)

	IF CurrPage = "" then CurrPage=1
	if page = "" then page = 1
	

	'### ���� ��� ���� ����
	set cMomoBonusCoinList = new ClsMomoCoin
	cMomoBonusCoinList.FPageSize = 20
	cMomoBonusCoinList.FCurrPage = page
	cMomoBonusCoinList.FUserID = userid
	cMomoBonusCoinList.FBonusCoinList
%>

<script language="javascript">
function goCoinEdit(idx)
{
	frame.document.location.href = "bonus_coin_give.asp?idx="+idx+"";
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    ���̵�:<input type="text" name="userid" value="<%=userid%>" size="10">
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
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

<table width="100%" align="center" cellpadding="3" cellspacing="2" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td><iframe id="frame" name="frame" src="bonus_coin_give.asp" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
</table>

<br>

<!-- ����Ʈ ���� -->
<table align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if cMomoBonusCoinList.FResultCount > 0 then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= cMomoBonusCoinList.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="center" width="50">�α�ID</td>
	    <td align="center" width="150">ȸ�����̵�</td>
	    <td align="center" width="100">Coin</td>
	    <td align="center" width="300">���ʽ� ���� ����</td>
	    <td align="center" width="150">�����</td>
	    <td align="center" width="60"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to cMomoBonusCoinList.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF">	
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fid %></td>
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fuserid %></td>
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fcoin %></td>
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fgubun %></td>
	    <td align="center"><%= cMomoBonusCoinList.FItemList(i).fregdate %></td>
		<td align="center"><input type="button" value="����" onClick="javascript:goCoinEdit('<%= cMomoBonusCoinList.FItemList(i).fid %>');"></td>
	</tr>
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if cMomoBonusCoinList.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= cMomoBonusCoinList.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + cMomoBonusCoinList.StartScrollPage to cMomoBonusCoinList.StartScrollPage + cMomoBonusCoinList.FScrollCount - 1 %>
				<% if (i > cMomoBonusCoinList.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(cMomoBonusCoinList.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if cMomoBonusCoinList.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set cMomoBonusCoinList = nothing	
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
