<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteBrandCls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
</head>
<body>
<%

dim xSiteId, gubun, idx

xSiteId = requestCheckvar(request("xSiteId"),32)
gubun = requestCheckvar(request("gubun"),32)
idx = requestCheckvar(request("idx"),32)

Dim oCxSiteBrand
set oCxSiteBrand = new CxSiteBrand

	if (idx = "") then
		oCxSiteBrand.FRectIdx   	= "-1"
		oCxSiteBrand.getXSiteBrandOne

		oCxSiteBrand.FOneItem.FxSiteId = xSiteId
		oCxSiteBrand.FOneItem.Fgubun = gubun
	else
		oCxSiteBrand.FRectIdx   	= idx
		oCxSiteBrand.getXSiteBrandOne
	end if

%>
<script language="javascript">
function jsSubmitIns() {
	var frm = document.frm;

	if(frm.xSiteId.value == "") {
		alert("���޸��� �����ϼ���.");
		frm.xSiteId.focus();
		return;
	}

	if(frm.gubun.value == "") {
		alert("������ �����ϼ���.");
		frm.gubun.focus();
		return;
	}

	if(frm.makerid.value == "") {
		alert("�귣�带 �Է��ϼ���.");
		frm.makerid.focus();
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "ins";
	frm.submit();
}

function jsSubmitDel() {
	var frm = document.frm;

	if (confirm("������ �����Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "del";
	frm.submit();
}

</script>

<b>���޸� �귣�����</b>

<form name="frm" action="xSiteBrandManage_process.asp" methd="post" style="margin:0px;" onSubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="80">���޸�</td>
	<td bgcolor="#FFFFFF" align="left"><% call drawSelectBoxXSiteOrderInputPartner("xSiteId", oCxSiteBrand.FOneItem.FxSiteId) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>����</td>
	<td bgcolor="#FFFFFF" align="left">
		<select class="select" name="gubun"  >
			<option value="" <%= chkIIF(oCxSiteBrand.FOneItem.Fgubun="", "selected","") %> >��ü</option>
	     	<option value="excoupon" <%= chkIIF(oCxSiteBrand.FOneItem.Fgubun="excoupon","selected","") %> >�������ܺ귣��</option>
     	</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�귣��</td>
	<td bgcolor="#FFFFFF" align="left"><% drawSelectBoxDesignerwithName "makerid", oCxSiteBrand.FOneItem.Fmakerid %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>�޸�</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea class="textarea" name="comment" cols="40" rows="4"><%= oCxSiteBrand.FOneItem.Fcomment %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitIns()">
		<% if (idx <> "") then %>
			&nbsp;&nbsp;
			<input type="button" class="button" value="����" onClick="jsSubmitDel()">
		<% end if %>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
