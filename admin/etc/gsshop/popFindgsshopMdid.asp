<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim ogsshop, i, page, isMapping

page		= request("page")
isMapping	= request("ismap")

If page = ""	Then page = 1

'// ��� ����
Set ogsshop = new CGSShop
	ogsshop.getgsshopMdidList
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

	// ī�װ� ����
	function fnSelmdid(dspNo,dspNm) {
	    opener.document.frmAct.mdid.value=dspNo;
		opener.document.getElementById("BrRow").style.display="";
		opener.document.getElementById("selBr").innerHTML="[" + dspNo + "] " + dspNm;
		self.close();
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
	<font color="red"><strong>gsshop MDID �˻�</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- �׼� -->
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=ogsshop.FtotalCount%></strong></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
	<td>GS�� MDID</td>
	<td>GS�� MD��</td>
</tr>
<% If ogsshop.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to ogsshop.FresultCount - 1
%>
<tr align="center" height="25" onClick="fnSelmdid('<%= ogsshop.FItemList(i).FMdid %>', '<%= ogsshop.FItemList(i).FMdname %>')" style="cursor:pointer" title="ī�װ� ����" bgcolor="#FFFFFF"">
	<td><%= ogsshop.FItemList(i).FMdid %></td>
	<td><%= ogsshop.FItemList(i).FMdname %></td>
</tr>
<%
		Next
	End If
%>
</table>
<iframe name="xLink" id="xLink" frameborder="1" width="10" height="10"></iframe>
<% Set ogsshop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
