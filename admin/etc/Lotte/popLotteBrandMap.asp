<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<%
	dim oLotte, i, page, srcDiv, srcKwd

	page		= request("page")
	srcDiv		= request("srcDiv")
	srcKwd		= request("srcKwd")
	if page="" then page=1
	if srcDiv="" then srcDiv="BNM"

	'// ��� ����
	Set oLotte = new cLotte
	oLotte.FPageSize = 20
	oLotte.FCurrPage = page
	oLotte.FRectSDiv = srcDiv
	oLotte.FRectKeyword = srcKwd
	oLotte.getLotteBrandList

%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value=pg;
		frm.submit();
	}

	// �˻�
	function serchItem() {
		frm.page.value=1;
		frm.submit();
	}

	// �ٹ����� �귣�� ��Ī �˾�
	function addLotteBrand(mkid) {
		var pFB = window.open("popFindLotteBrand.asp?mkid="+mkid,"popFindBrand","width=400,height=300,scrollbars=yes,resizable=yes");
		pFB.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>�Ե����� �귣�� ����</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
<form name="frm" method="GET" style="margin:0px;">
<input type="hidden" name="page" value="<%=page%>">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding-top:5px;">
		�˻����� :
		<select name="srcDiv" class="select">
			<option value="LCD" <%=chkIIF(srcDiv="LCD","selected","")%>>�Ե����� �ڵ�</option>
			<option value="TCD" <%=chkIIF(srcDiv="TCD","selected","")%>>�ٹ����� �ڵ�</option>
			<option value="BNM" <%=chkIIF(srcDiv="BNM","selected","")%>>�귣���</option>
		</select> /
		�˻��� :
		<input type="text" name="srcKwd" size="15" value="<%=srcKwd%>" class="text"> &nbsp;
		<input id="btnRefresh" type="button" class="button" value="�˻�" onclick="serchItem()">
	</td>
</tr>
<tr>
	<td align="right" style="padding-top:5px;"><input id="btnRefresh" type="button" class="button" value="+�귣�� ��Ī" onclick="addLotteBrand('')"></td>
</tr>
</table>
</form>
<p>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oLotte.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="2">�Ե����� �귣��</td>
	<td colspan="2">�ٹ����� �귣��</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�ڵ�</td>
	<td>�귣���</td>
	<td>���̵�</td>
	<td>�귣���</td>
</tr>
<% if oLotte.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	else
		for i=0 to oLotte.FresultCount-1
%>
<tr align="center" height="25" bgcolor="<%=chkIIF(oLotte.FItemList(i).FisUsing="Y","#FFFFFF","#DDDDDD")%>" title="��Ī ����" onClick="addLotteBrand('<%= oLotte.FItemList(i).FTenMakerid %>')" style="cursor:pointer">
	<td><%= oLotte.FItemList(i).FlotteBrandCd %></td>
	<td><%= oLotte.FItemList(i).FlotteBrandName %></td>
	<td><%= oLotte.FItemList(i).FTenMakerid %></td>
	<td><%= oLotte.FItemList(i).FTenBrandName %></td>
</tr>
<%
		next
	end if
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% if oLotte.HasPreScroll then %>
		<a href="javascript:goPage('<%= oLotte.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oLotte.StartScrollPage to oLotte.FScrollCount + oLotte.StartScrollPage - 1 %>
			<% if i>oLotte.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oLotte.HasNextScroll then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oLotte = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->