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
		alert("제휴몰을 선택하세요.");
		frm.xSiteId.focus();
		return;
	}

	if(frm.gubun.value == "") {
		alert("구분을 선택하세요.");
		frm.gubun.focus();
		return;
	}

	if(frm.makerid.value == "") {
		alert("브랜드를 입력하세요.");
		frm.makerid.focus();
		return;
	}

	if (confirm("저장하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "ins";
	frm.submit();
}

function jsSubmitDel() {
	var frm = document.frm;

	if (confirm("정말로 삭제하시겠습니까?") != true) {
		return;
	}

	frm.mode.value = "del";
	frm.submit();
}

</script>

<b>제휴몰 브랜드관리</b>

<form name="frm" action="xSiteBrandManage_process.asp" methd="post" style="margin:0px;" onSubmit="return false;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="idx" value="<%= idx %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="80">제휴몰</td>
	<td bgcolor="#FFFFFF" align="left"><% call drawSelectBoxXSiteOrderInputPartner("xSiteId", oCxSiteBrand.FOneItem.FxSiteId) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>구분</td>
	<td bgcolor="#FFFFFF" align="left">
		<select class="select" name="gubun"  >
			<option value="" <%= chkIIF(oCxSiteBrand.FOneItem.Fgubun="", "selected","") %> >전체</option>
	     	<option value="excoupon" <%= chkIIF(oCxSiteBrand.FOneItem.Fgubun="excoupon","selected","") %> >쿠폰제외브랜드</option>
     	</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>브랜드</td>
	<td bgcolor="#FFFFFF" align="left"><% drawSelectBoxDesignerwithName "makerid", oCxSiteBrand.FOneItem.Fmakerid %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>메모</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea class="textarea" name="comment" cols="40" rows="4"><%= oCxSiteBrand.FOneItem.Fcomment %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="button" class="button" value="저장하기" onClick="jsSubmitIns()">
		<% if (idx <> "") then %>
			&nbsp;&nbsp;
			<input type="button" class="button" value="삭제" onClick="jsSubmitDel()">
		<% end if %>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
