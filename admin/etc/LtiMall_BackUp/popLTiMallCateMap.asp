<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/lotteiMallcls.asp"-->
<%
Dim oiMall, i
Dim cdl, cdm, cds, dispNo, dispNm, dispFull
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")
dispNo	= request("dspNo")

If cdl = "" Then
	Call Alert_Close("ī�װ� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ� ���� ����
Set oiMall = new CLotteiMall
	oiMall.FPageSize = 20
	oiMall.FCurrPage = 1
	oiMall.FRectCDL = cdl
	oiMall.FRectCDM = cdm
	oiMall.FRectCDS = cds
	oiMall.FRectDspNo = dispNo
	oiMall.getTenLotteimallCateList

If oiMall.FResultCount <= 0 then
	Call Alert_Close("�ش� ī�װ� ������ �����ϴ�.")
	dbget.Close: Response.End
End If

dispNo = oiMall.FItemList(0).FDispNo
dispNm = oiMall.FItemList(0).FDispNm
dispFull = oiMall.FItemList(0).FDispLrgNm
If Not(oiMall.FItemList(0).FDispMidNm="" or isNull(oiMall.FItemList(0).FDispMidNm)) then dispFull = dispFull & " > " & oiMall.FItemList(0).FDispMidNm
If Not(oiMall.FItemList(0).FDispSmlNm="" or isNull(oiMall.FItemList(0).FDispSmlNm)) then dispFull = dispFull & " > " & oiMall.FItemList(0).FDispSmlNm
If Not(oiMall.FItemList(0).FDispThnNm="" or isNull(oiMall.FItemList(0).FDispThnNm)) then dispFull = dispFull & " > " & oiMall.FItemList(0).FDispThnNm
%>
<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frm;
		if(frm.dspNo.value=="") {
			alert("��Ī�� �Ե����̸� ī�װ��� ���� ���ּ���.");
			return;
		}
		if(confirm("�����Ͻ� ī�װ��� ��Ī�Ͻðڽ��ϱ�?")) {
			frm.mode.value="save";
			frm.action="procLtiMallCateMapping.asp";
			frm.submit();
		}
	}

	// ��Ī���� ����
	function fnDelForm() {
		var frm = document.frm;
		<% if Not(dispNo="" or isNull(dispNo)) then %>
		if(confirm("���� ��Ī�� ī�װ��� �������� �Ͻðڽ��ϱ�?\n\n�� ��ǰ �Ǵ� ī�װ��� �����Ǵ� ���� �ƴϸ�, ����� ������ �����˴ϴ�.")) {
			frm.mode.value="del";
			frm.action="procLtiMallCateMapping.asp";
			frm.submit();
		}
		<% else %>
			alert("��Ī�� �Ե����̸� ī�װ��� �����ϴ�.");
			return;
		<% end if %>
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// �Ե����̸� ī�װ� �˻�
	function fnSearchLotteCate() {
		var pFCL = window.open("","popLotteCate","width=700,height=600,scrollbars=yes,resizable=yes");
		pFCL.focus();
		srcFrm.target="popLotteCate";
		srcFrm.action="popFindLtiMallCate.asp";
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
	<font color="red"><strong>�Ե����̸� ī�װ� ��Ī</strong></font></td>
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
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� ī�װ� ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��з�</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=oiMall.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ߺз�</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=oiMall.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�Һз�</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=oiMall.FItemList(0).FtenCDSName%></td>
</tr>
</table>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �Ե����̸� ���� ī�װ� ��Ī ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td id="brTT" width="80" align="center" bgcolor="#DDDDFF" <%=chkIIF(Not(dispNo="" or isNull(dispNo)),"rowspan=2","")%>>�˻�</td>
	<td bgcolor="#FFFFFF">
		ī�װ��� <input type="text" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchLotteCate()">
	</td>
</tr>
<tr id="BrRow" style="display:<%=chkIIF(Not(dispNo="" or isNull(dispNo)),"","none")%>">
	<td bgcolor="#F2F2F2">���� : <b><span id="selBr">[<%=dispNo%>] <%=dispNm%></span></b>
	<% if Not(dispFull="" or isNull(dispFull)) then Response.Write "<br>" & dispFull %>
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
<form name="frm" method="GET" target="xLink" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="dspNo" value="<%=dispNo%>">
<input type="hidden" name="mode" value="save">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
</p>
<% Set oiMall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
