<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall2/cjmallitemcls.asp"-->
<%
Dim ocjmall, i
Dim cdl, cdm, cds, dispNo '', dispNm, dispFull
Dim infodiv, infodivnm, mode
mode	= request("mode")
cdl		= request("cdl")
cdm		= request("cdm")
cds		= request("cds")
infodiv	= request("infodiv")
dispNo	= request("dspNo")

Select Case infodiv
	Case "01"	infodivnm = "�Ƿ�"
	Case "02"	infodivnm = "����/�Ź�"
	Case "03"	infodivnm = "����"
	Case "04"	infodivnm = "�м���ȭ(����/��Ʈ/�׼�����)"
	Case "05"	infodivnm = "ħ����/Ŀư"
	Case "06"	infodivnm = "����(ħ��/����/��ũ��/DIY��ǰ)"
	Case "07"	infodivnm = "������(TV��)"
	Case "08"	infodivnm = "������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ�����)"
	Case "09"	infodivnm = "��������(������/��ǳ��)"
	Case "10"	infodivnm = "�繫����(��ǻ��/��Ʈ��/������)"
	Case "11"	infodivnm = "���б��(������ī�޶�/ķ�ڴ�)"
	Case "12"	infodivnm = "��������(MP3/���ڻ��� ��)"
	Case "13"	infodivnm = "�޴���"
	Case "14"	infodivnm = "������̼�"
	Case "15"	infodivnm = "�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)"
	Case "16"	infodivnm = "�Ƿ���"
	Case "17"	infodivnm = "�ֹ��ǰ"
	Case "18"	infodivnm = "ȭ��ǰ"
	Case "19"	infodivnm = "�ͱݼ�/����/�ð��"
	Case "20"	infodivnm = "��ǰ(����깰)"
	Case "21"	infodivnm = "������ǰ"
	Case "22"	infodivnm = "�ǰ���ɽ�ǰ"
	Case "23"	infodivnm = "�����ƿ�ǰ"
	Case "24"	infodivnm = "�Ǳ�"
	Case "25"	infodivnm = "��������ǰ"
	Case "26"	infodivnm = "����"
	Case "27"	infodivnm = "ȣ��/��� ����"
	Case "28"	infodivnm = "������Ű��"
	Case "29"	infodivnm = "�װ���"
	Case "30"	infodivnm = "�ڵ��� �뿩 ����(����ī)"
	Case "31"	infodivnm = "��ǰ�뿩 ����(������, ��, ����û���� ��)"
	Case "32"	infodivnm = "��ǰ�뿩 ����(����, ���ƿ�ǰ, ����ǰ ��)"
	Case "33"	infodivnm = "������ ������(����, ����, ���ͳݰ��� ��)"
	Case "34"	infodivnm = "��ǰ��/����"
	Case "35"	infodivnm = "��Ÿ"
End Select


If cdl = "" Then
	Call Alert_Close("ī�װ� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ� ���� ����
Set ocjmall = new CCjmall
	ocjmall.FRectCDL = cdl
	ocjmall.FRectCDM = cdm
	ocjmall.FRectCDS = cds
	ocjmall.Finfodiv = infodiv
	ocjmall.getTencjmallOneprdDiv

'If ocjmall.FResultCount <= 0 Then
'	Call Alert_Close("�ش� ī�װ� ������ �����ϴ�.")
'	dbget.Close: Response.End
'End If
%>
<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.dspNo.value=="") {
			alert("��Ī�� cjmall ��ǰ�з��� �������ּ���.");
			return;
		}

		if(confirm("�����Ͻ� ��ǰ�з��� ��Ī�Ͻðڽ��ϱ�?")) {
			frm.mode.value="saveCate";
			frm.action="proccjmall2.asp";
			frm.submit();
		}
	}

    function fnDelForm(iDspNo) {
		var frm = document.frmAct;
		if (iDspNo=="") {
		    alert("������ cjmall ��ǰ�з��� �������ּ���.");
			return;
		}

		if(confirm("���� ��Ī�� ��ǰ�з��� �������� �Ͻðڽ��ϱ�?\n\n�� ��ǰ�з��� �����Ǵ� ���� �ƴϸ�, ����� ������ �����˴ϴ�.")) {
			frm.mode.value="delPrddiv";
			frm.dspNo.value=iDspNo;
			frm.action="proccjmall2.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// cjmall ��ǰ�з� �˻�
	function fnSearchCJPrddiv(disptpcd, cddnm) {
		var pFCL = window.open("popFindcjmallPrddiv.asp?infodiv="+disptpcd+"&cdd_NAME="+cddnm,"popcjmallPrddiv","width=900,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
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
	<font color="red"><strong>cjmall ��ǰ�з� ��Ī</strong></font></td>
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
	<td width="80" align="center" bgcolor="#DDDDFF">ǰ������</td>
	<td bgcolor="#FFFFFF">[<%=infodiv%>] <%=infodivnm%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��з�</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=ocjmall.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ߺз�</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=ocjmall.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�Һз�</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=ocjmall.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> cjmall ��ǰ�з� ��Ī ����</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% If mode <> "U" Then %>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >�˻�</td>
	<td bgcolor="#FFFFFF">
		���з��� <input type="text" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchCJPrddiv('<%=infodiv%>',document.srcFrm.srcKwd.value)">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">�߰� : <b><span id="selBr"></span></b></td>
</tr>
<% End If %>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= ocjmall.FResultCount + 1 %>" >��ϵ�<br>��ǰ�з�</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% If Not IsNULL(ocjmall.FItemList(0).FCddKey) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr">[<%=ocjmall.FItemList(0).FCddKey%>] <%=ocjmall.FItemList(0).Fcdd_Name%></span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%=ocjmall.FItemList(0).FCddKey%>')" style="cursor:pointer" align="absmiddle">
    </td>
</tr>
<% End If %>
</table>
</form>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr valign="top" height="28">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="left"></td>
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
<form name="frmAct" method="POST" target="xLink" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="dspNo" value="">
<input type="hidden" name="mode" value="saveCate">
<input type="hidden" name="infodiv" value="<%=infodiv%>">
<input type="hidden" name="CdmKey" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="110" height="110"></iframe>
</p>
<% Set ocjmall = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
