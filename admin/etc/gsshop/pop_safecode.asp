<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/gsshop/gsshopItemcls.asp"-->
<%
Dim itemid, oGSShop, mode
itemid = request("itemid")
Set oGSShop = new CGSShop
	oGSShop.FRectItemid = itemid
	oGSShop.getgsshopSafeCodeList

	If oGSShop.FResultCount < 1 Then
		Call Alert_Close("������������ ��ϵ� ��ǰ�� �ƴմϴ�.")
		dbget.Close: Response.End
	End If

	If oGSShop.FItemList(0).FSafeCertGbnCd <> "" Then
		mode = "U"
	Else
		mode = "I"
	End If
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function fnSaveForm() {
	var frm = document.frmAct;
	if(frm.safeCertOrgCd.value=="") {
		alert("��������� �Է��ϼ���");
		frm.safeCertOrgCd.focus();
		return;
	}
	if(frm.safeCertModelNm.value=="") {
		alert("�����𵨸��� �Է��ϼ���");
		frm.safeCertModelNm.focus();
		return;
	}
	if(frm.safeCertNo.value=="") {
		alert("������ȣ�� �Է��ϼ���");
		frm.safeCertNo.focus();
		return;
	}
	if(frm.safeCertDt.value=="") {
		alert("�������� �Է��ϼ���");
		return;
	}
	frm.target = "xLink";
	frm.action = "/admin/etc/gsshop/proc_safecode.asp"
	frm.submit();	
}	
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
	<font color="red"><strong>GSShop �ʼ� �������� ��Ī</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� �������� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF"><%= oGSShop.FItemList(0).FItemid %></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">������������</td>
	<td bgcolor="#FFFFFF">
	<%
		Select Case oGSShop.FItemList(0).FSafetyDiv
			Case "10"	response.write "������������(KC��ũ)"
			Case "20"	response.write "�����ǰ ��������"
			Case "30"	response.write "KPS �������� ǥ��"
			Case "40"	response.write "KPS �������� Ȯ�� ǥ��"
			Case "50"	response.write "KPS ��� ��ȣ���� ǥ��"
		End Select
	%>
	</td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">������ȣ</td>
	<td bgcolor="#FFFFFF"><%=oGSShop.FItemList(0).FSafetyNum%></td>
</tr>
</table>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> GSShop �������� ��Ī</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="frmAct" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="itemid" value="<%=oGSShop.FItemList(0).FItemid%>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">������������</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="safeCertGbnCd">
			<option value="1" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "1","selected","")%> >�����������</option>
			<option value="2" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "2","selected","")%> >����ǰ��������</option>
			<option value="3" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "3","selected","")%> >����ǰ��������Ȯ�ι�ȣ</option>
			<option value="4" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "4","selected","")%> >�����ǰ��������Ȯ��</option>
			<option value="5" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertGbnCd = "5","selected","")%> >�����ű���������</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">�������</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="safeCertOrgCd">
			<option value="101" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "101","selected","")%> >�ѱ� �������ڽ��� ������</option>
			<option value="102" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "102","selected","")%> >������ �����</option>
			<option value="103" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "103","selected","")%> >�ѱ� �����Ŀ�����</option>
			<option value="104" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "104","selected","")%> >�ѱ�ǥ����ȸ</option>
			<option value="201" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "201","selected","")%> >�ѱ� ��Ȱȯ�� ���迬����</option>
			<option value="202" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "202","selected","")%> >�ѱ� �Ƿ����� ������</option>
			<option value="203" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "203","selected","")%> >�ѱ� ȭ�н��� ������</option>
			<option value="204" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "204","selected","")%> >�ѱ� �����ȭ ���迬����</option>
			<option value="205" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "205","selected","")%> >�ѱ� �������� ���迬����</option>
			<option value="206" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "206","selected","")%> >�ѱ� ������ ���迬����</option>
			<option value="207" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "207","selected","")%> >������ �����</option>
			<option value="208" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "208","selected","")%> >�ѱ� �ϱ����� ��������</option>
			<option value="301" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "301","selected","")%> >�ѱ� ��Ȱȯ�� ���迬����</option>
			<option value="302" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "302","selected","")%> >�ѱ� �Ƿ����� ������</option>
			<option value="303" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "303","selected","")%> >�ѱ� ȭ�н��� ������</option>
			<option value="304" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "304","selected","")%> >�ѱ� �����ȭ ���迬����</option>
			<option value="305" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "305","selected","")%> >�ѱ� �������� ���迬����</option>
			<option value="306" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "306","selected","")%> >�ѱ� ������ ���迬����</option>
			<option value="307" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "307","selected","")%> >������ �����</option>
			<option value="308" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "308","selected","")%> >�ѱ� �ϱ����� ��������</option>
			<option value="401" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "401","selected","")%> >�ѱ������������</option>
			<option value="402" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "402","selected","")%> >�ѱ�����������ڽ��迬����</option>
			<option value="403" <%= ChkIIF(oGSShop.FItemList(0).FSafeCertOrgCd = "403","selected","")%> >�ѱ�ȭ�����ս��迬����</option>
		</select>
	</td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">�����𵨸�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="safeCertModelNm" maxlength="100" value="<%=oGSShop.FItemList(0).FSafeCertModelNm%>" size="50"></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">������ȣ</td>
	<td bgcolor="#FFFFFF"><input type="text" name="safeCertNo" maxlength="30" value="<%=oGSShop.FItemList(0).FSafeCertNo%>" size="30"></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF">������</td>
	<td bgcolor="#FFFFFF">
        <input id="safeCertDt" name="safeCertDt" value="<%=oGSShop.FItemList(0).FSafeCertDt%>" class="text" size="10" maxlength="10" readonly />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="safeCertDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "safeCertDt", trigger    : "safeCertDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"><a href="http://www.safetykorea.kr/search/search_pop.html?authNum=<%=oGSShop.FItemList(0).FSafetyNum%>" target="_blank"><font color="GREEN"><strong>��ȸ�ϱ�</strong></font></a></td>
    <td valign="bottom" align="right">
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
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
<% Set oGSShop = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
