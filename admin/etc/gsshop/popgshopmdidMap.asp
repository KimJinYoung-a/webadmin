<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim ogsshop, i
Dim catekey, reged
catekey = request("catekey")
reged	= request("reged")
If catekey = "" Then
	Call Alert_Close("ī�װ� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ� ���� ����
Set ogsshop = new CGSShop
	ogsshop.FPageSize 		= 20
	ogsshop.FCurrPage		= 1
	ogsshop.FRectCatekey	= catekey
	ogsshop.getTengsshopMdidList

If ogsshop.FResultCount <= 0 Then
	Call Alert_Close("�ش� MDID ������ �����ϴ�.")
	dbget.Close: Response.End
End If
%>
<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.mdid.value=="") {
			alert("��Ī�� MDID�� �������ּ���.");
			return;
		}

		if(confirm("�����Ͻ� MDID�� ��Ī�Ͻðڽ��ϱ�?")) {
			frm.mode.value="saveMD";
			frm.action="procgsshop2.asp";
			frm.submit();
		}
	}

    function fnDelForm(iDspNo,mdid) {
		var frm = document.frmAct;
		if (iDspNo=="") {
		    alert("������ MDID�� �������ּ���.");
			return;
		}

		if(confirm("���� MDID�� �������� �Ͻðڽ��ϱ�?")) {
			frm.mode.value="delMdid";
			frm.catekey.value=iDspNo;
			frm.mdid.value=mdid;
			frm.action="procgsshop2.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// gsshop ī�װ� �˻�
	function fnSearchMdid(disptpcd) {
		var pFCL = window.open("","popgsshopMdid","width=900,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
		srcFrm.target="popgsshopMdid";
		srcFrm.action="popFindgsshopMdid.asp";
		srcFrm.submit();
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
	<font color="red"><strong>gsshop MDID ��Ī</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� ���� ī�װ� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��</td>
	<td bgcolor="#FFFFFF"><%=ogsshop.FItemList(0).FL_NAME %></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��</td>
	<td bgcolor="#FFFFFF"><%=ogsshop.FItemList(0).FM_NAME%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��</td>
	<td bgcolor="#FFFFFF"><%=ogsshop.FItemList(0).FS_NAME%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��</td>
	<td bgcolor="#FFFFFF"><%=ogsshop.FItemList(0).FD_NAME%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ڵ�</td>
	<td bgcolor="#FFFFFF"><%=ogsshop.FItemList(0).FCatekey %></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> gsshop MDID ��Ī ���� </td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >�˻�</td>
	<td bgcolor="#FFFFFF">
		MDID <input type="text" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchMdid()">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">�߰� : <b><span id="selBr"></span></b></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= ogsshop.FResultCount + 1 %>" >��ϵ�<br>MDID</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to ogsshop.FResultCount - 1 %>
<% If ogsshop.FItemList(i).FMdid <> "" Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=ogsshop.FItemList(i).getDispGubunNm%> [<%=ogsshop.FItemList(i).FMdid%>] <%=ogsshop.FItemList(i).FMdname%></span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%=ogsshop.FItemList(i).FCatekey %>', '<%=ogsshop.FItemList(i).FMdid %>')" style="cursor:pointer" align="absmiddle">
    </td>
</tr>
<% End If %>
<% Next %>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
    <td valign="bottom" align="right">
		<img src="http://testwebadmin.10x10.co.kr/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<% If reged <> "Y" Then %>
		<img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle">
		<% End If %>
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
<form name="frmAct" method="POST" target="xLink" style="margin:0px;">
<input type="hidden" name="mode" value="saveMD">
<input type="hidden" name="catekey" value="<%= catekey %>">
<input type="hidden" name="mdid" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
</p>
<% Set ogsshop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
