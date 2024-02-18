<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/lotteitemcls.asp"-->
<%
	dim oLotte
	dim TenMakerid, lotteBrandCd, lotteBrandName

	TenMakerid		= request("mkid")

	if TenMakerid<>"" then
		'// ��� ����
		Set oLotte = new cLotte
		oLotte.FPageSize = 20
		oLotte.FCurrPage = 1
		oLotte.FRectMakerid = TenMakerid
		oLotte.getLotteBrandList
		if oLotte.FResultCount>0 then
			lotteBrandCd = oLotte.FItemList(0).FlotteBrandCd
			lotteBrandName = oLotte.FItemList(0).FlotteBrandName
		end if
		Set oLotte = Nothing
	end if
%>
<script language="javascript">
<!--
	// �Ե����� �귣�� �˻�
	function fnSearchLotteBrand() {
		if(!fsrch.brnNm.value) {
			alert("�˻�� �Է����ּ���.(ex.�귣���)");
			fsrch.brnNm.focus();
			return;
		}
		var pFBL = window.open("","popLotteBrand","width=400,height=500,scrollbars=yes,resizable=yes");
		pFBL.focus();
		fsrch.target="popLotteBrand";
		fsrch.action="actFindLotteBrand.asp";
		fsrch.submit();
	}

	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frm;
		if(frm.TenMakerid.value=="") {
			alert("��Ī�� �ٹ����� �귣�带 �������ּ���.");
			frm.TenMakerid.focus();
			return;
		}

		if(frm.lotteBrandCd.value=="") {
			alert("��Ī�� �Ե����� �귣�带 �������ּ���.");
			return;
		}

		if(confirm("�����Ͻ� �귣�带 ���� ��Ī�Ͻðڽ��ϱ�?")) {
			frm.mode.value="save";
			frm.action="procLotteBrandMap.asp";
			frm.submit();
		}
	}

	// ��Ī���� ����
	function fnDelForm() {
		var frm = document.frm;
		<% if TenMakerid="" then %>
		alert("��Ī�� �귣�尡 �����ϴ�.");
		return;
		<% else %>
		if(confirm("�귣�带 ��Ī �����Ͻðڽ��ϱ�?")) {
			frm.mode.value="del";
			frm.action="procLotteBrandMap.asp";
			frm.submit();
		}
		<% end if %>
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
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
	<font color="red"><strong>�Ե����� �귣�� ��Ī</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� �귣�� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="frm" method="get" action="" target="xLink" style="margin:0px;">
<input type="hidden" name="mode" value="save">
<input type="hidden" name="lotteBrandCd" value="<%=lotteBrandCd%>">
<input type="hidden" name="lotteBrandNm" value="<%=lotteBrandName%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�귣��ID</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "TenMakerid",TenMakerid %></td>
</tr>
</table>
</form>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �Ե����� �귣�� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="fsrch" method="get" action="actFindLotteBrand.asp" style="margin:0px;" onsubmit="return false;">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td id="brTT" <%=chkIIF(TenMakerid<>"","rowspan=2","")%> width="80" align="center" bgcolor="#DDDDFF">�귣�� �˻�</td>
	<td bgcolor="#FFFFFF">�귣��� <input type="text" name="brnNm"  size="12" class="text">
		<input type="button" value="�˻�" class="button" onclick="fnSearchLotteBrand()">
	</td>
</tr>
<tr id="BrRow" height="25" <%=chkIIF(TenMakerid<>"","","style='display:none;'")%>>
	<td bgcolor="#FFFFFF">
		���� �귣�� : <span id="selBr">[<%=lotteBrandCd%>]<%=lotteBrandName%></span>
	</td>
</tr>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"><img src="http://testwebadmin.10x10.co.kr/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm()" style="cursor:pointer" align="absmiddle"></td>
    <td valign="bottom" align="right">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="bottom" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
