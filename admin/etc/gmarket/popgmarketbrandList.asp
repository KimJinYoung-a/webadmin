<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gmarket/gmarketcls.asp"-->
<%
Dim oGmarket, i, page, makerid, isbrandcd
page			= request("page")
makerid			= request("makerid")
isbrandcd		= request("isbrandcd")
If page = ""	Then page = 1

'// ��� ����
Set oGmarket = new CGmarket
	oGmarket.FPageSize 			= 20
	oGmarket.FCurrPage			= page
	oGmarket.FRectIsbrandcd		= isbrandcd
	oGmarket.FRectMakerid		= makerid
	oGmarket.getTenGmarketBrandList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg) {
		frm.page.value = pg;
		frm.submit();
	}

	// �˻�
	function serchItem() {
		frm.page.value = 1;
		frm.submit();
	}

	// Gmarket �귣���ڵ� ��Ī �˾�
	function popBrandMap(makerid) {
		var pCM = window.open("popgmarketBrandMap.asp?makerid="+makerid,"popBrandMap","width=600,height=500,scrollbars=yes,resizable=yes");
		pCM.focus();
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
	<font color="red"><strong>Gmarket �귣���ڵ� ����</strong></font></td>
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
		<br>
		�� �� �� : <% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
		�귣���ڵ� ��Ī���� :
		<select name="isbrandcd" class="select">
			<option value="">��ü</option>
			<option value="Y" <%=chkIIF(isbrandcd="Y","selected","")%>>��Ī�Ϸ�</option>
			<option value="N" <%=chkIIF(isbrandcd="N","selected","")%>>�̸�Ī</option>
		</select>
	</td>
	<td width="55" align="right" style="padding-top:5px;">
		<input id="btnRefresh" type="button" class="button" value="�˻�" onclick="serchItem()" style="width:50px;height:40px;">
	</td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oGmarket.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#E8E8FF">
	<td colspan="3">�ٹ����� �귣��</td>
	<td>Gmarket �ڵ�</td>
</tr>
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>�귣��ID</td>
	<td>�귣���(�ѱ�)</td>
	<td>�귣���(����)</td>
	<td>�귣���ڵ�</td>
</tr>
<% If oGmarket.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oGmarket.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="<%= CHKIIF(oGmarket.FItemList(i).FBrandcode <> "0" ,"#FFFFFF","#CCCCCC") %>">
	<td><%= oGmarket.FItemList(i).FUserid %></td>
	<td><%= oGmarket.FItemList(i).FSocname_kor %></td>
	<td><%= oGmarket.FItemList(i).FSocname %></td>
	<% If oGmarket.FItemList(i).FBrandcode="0" Then %>
	<td ><input type="button" class="button" value="Gmarket ��Ī" onClick="popBrandMap('<%= oGmarket.FItemList(i).FUserid %>')"></td>
	<% Else %>
	<td style="cursor:pointer;" onclick="popBrandMap('<%= oGmarket.FItemList(i).FUserid %>')"><%= oGmarket.FItemList(i).FBrandcode %>&nbsp;[<%= oGmarket.FItemList(i).FMakername %>] </td>
	<% End If %>
</tr>
<%
		Next
	End If
%>
</table>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
		<% If oGmarket.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oGmarket.StartScrollPage-1 %>')">[pre]</a>
		<% Else %>
			[pre]
		<% End If %>
		<% For i = 0 + oGmarket.StartScrollPage to oGmarket.FScrollCount + oGmarket.StartScrollPage - 1 %>
			<% If i > oGmarket.FTotalpage Then Exit For %>
			<% If CStr(page)=CStr(i) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
			<% End If %>
		<% Next %>
		<% If oGmarket.HasNextScroll Then %>
			<a href="javascript:goPage('<%= i %>')">[next]</a>
		<% Else %>
			[next]
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<% Set oGmarket = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->