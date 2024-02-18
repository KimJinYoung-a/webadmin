<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/skstoa/skstoaCls.asp"-->
<%
Dim oSkstoa, i
Dim cdl, cdm, cds, lgroup, mgroup, sgroup, dgroup
cdl		= requestCheckVar(request("cdl"),3)
cdm		= requestCheckVar(request("cdm"),3)
cds		= requestCheckVar(request("cds"),3)
lgroup	= requestCheckVar(request("lgroup"),10)
mgroup	= requestCheckVar(request("mgroup"),10)
sgroup	= requestCheckVar(request("sgroup"),10)
dgroup	= requestCheckVar(request("dgroup"),10)

If cdl = "" Then
	Call Alert_Close("ī�װ� �ڵ尡 �����ϴ�.")
	dbget.Close: Response.End
End IF

'// ī�װ� ���� ����
Set oSkstoa = new CSkstoa
	oSkstoa.FPageSize = 20
	oSkstoa.FCurrPage = 1
	oSkstoa.FRectCDL = cdl
	oSkstoa.FRectCDM = cdm
	oSkstoa.FRectCDS = cds
	oSkstoa.getTenSkstoaCateList

If oSkstoa.FResultCount <= 0 Then
	Call Alert_Close("�ش� ī�װ� ������ �����ϴ�.")
	dbget.Close: Response.End
End If
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
<!--
	// ��Ī �����ϱ�
	function fnSaveForm() {
		var frm = document.frmAct;

		if(frm.lgroup.value=="") {
			alert("��Ī�� Skstoa ī�װ��� �������ּ���.");
			return;
		}

		if(confirm("�����Ͻ� ī�װ��� ��Ī�Ͻðڽ��ϱ�?")) {
			frm.mode.value="saveCate";
			frm.action="procskstoa.asp";
			frm.submit();
		}
	}

    function fnDelForm(cdl, cdm, cds) {
		var frm = document.frmAct;
		if (cdl=="") {
		    alert("������ Skstoa ī�װ��� �������ּ���.");
			return;
		}

		if(confirm("���� ��Ī�� ī�װ��� �������� �Ͻðڽ��ϱ�?\n\n�� ��ǰ �Ǵ� ī�װ��� �����Ǵ� ���� �ƴϸ�, ����� ������ �����˴ϴ�.")) {
			frm.mode.value="delCate";
			frm.cdl.value=cdl;
			frm.cdm.value=cdm;
			frm.cds.value=cds;
			frm.action="procskstoa.asp";
			frm.submit();
		}
	}

	// â�ݱ�
	function fnCancel() {
		if(confirm("�۾��� ����ϰ� â�� �����ðڽ��ϱ�?")) {
			self.close();
		}
	}

	// Skstoa ī�װ� �˻�
	function fnSearchSkstoaCate() {
		var kwd;
		kwd = document.getElementById("srcKwd").value;
		var pFCL = window.open("popFindskstoaCate.asp?srcKwd="+kwd,"popSkstoaCate","width=1000,height=700,scrollbars=yes,resizable=yes");
		pFCL.focus();
	}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr valign="top">
	<td><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>Skstoa ī�װ� ��Ī</strong></font></td>
</tr>
</table>
<p>
<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �ٹ����� ī�װ� ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">��з�</td>
	<td bgcolor="#FFFFFF">[<%=cdl%>] <%=oSkstoa.FItemList(0).FtenCDLName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�ߺз�</td>
	<td bgcolor="#FFFFFF">[<%=cdm%>] <%=oSkstoa.FItemList(0).FtenCDMName%></td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#DDDDFF">�Һз�</td>
	<td bgcolor="#FFFFFF">[<%=cds%>] <%=oSkstoa.FItemList(0).FtenCDSName%></td>
</tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="left"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> Skstoa ī�װ� ��Ī ����</td>
</tr>
</table>
<!-- ǥ �߰��� ��-->
<form name="srcFrm" method="GET" onsubmit="fnSearchSkstoaCate();return false;" style="margin:0px;">
<input type="hidden" name="srcDiv" value="CNM">
<input type="hidden" name="disptpcd" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="2" >�˻�</td>
	<td bgcolor="#FFFFFF">
		ī�װ��� <input type="text" id="srcKwd" name="srcKwd" class="text">
		<input type="button" value="�˻�" class="button" onClick="fnSearchSkstoaCate();">
	</td>
</tr>
<tr id="BrRow" style="display:">
	<td bgcolor="#F2F2F2">�߰� : <b><span id="selBr"></span></b></td>
</tr>
<tr height="25">
    <td id="brTT" width="80" align="center" bgcolor="#DDDDFF" rowspan="<%= oSkstoa.FResultCount + 1 %>" >��ϵ�<br>ī�װ�</td>
	<td bgcolor="#FFFFFF" height="1"></td>
</tr>
<% For i = 0 to oSkstoa.FResultCount - 1 %>
<% If Not IsNULL(oSkstoa.FItemList(i).FLgroup) Then %>
<tr>
    <td bgcolor="#F2F2F2"><b><span id="selBr"><%=oSkstoa.FItemList(i).FLastDepthNm%> [<%= oSkstoa.FItemList(i).FLgroup & oSkstoa.FItemList(i).FMgroup & oSkstoa.FItemList(i).FSgroup & oSkstoa.FItemList(i).FDgroup %>]</span></b>
    &nbsp;&nbsp;&nbsp;&nbsp;
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
		<img src="/images/icon_cancel.gif" width="45" height="20" border="0" onclick="fnCancel()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<img src="/images/icon_save.gif" width="45" height="20" border="0" onclick="fnSaveForm()" style="cursor:pointer" align="absmiddle"> &nbsp;&nbsp;&nbsp;
		<% If lgroup <> "" Then %>
		<img src="/images/icon_delete.gif" width="45" height="20" border="0" onclick="fnDelForm('<%= cdl %>', '<%= cdm %>', '<%= cds %>');" style="cursor:pointer" align="absmiddle">
		<% End If %>
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<form name="frmAct" method="POST" style="margin:0px;">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="lgroup" value="<%= lgroup %>">
<input type="hidden" name="mgroup" value="<%= mgroup %>">
<input type="hidden" name="sgroup" value="<%= sgroup %>">
<input type="hidden" name="dgroup" value="<%= dgroup %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="categbn" value="cate">
</form>
<iframe name="xLink" id="xLink" frameborder="0" width="1110" height="110"></iframe>
</p>
<% Set oSkstoa = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
