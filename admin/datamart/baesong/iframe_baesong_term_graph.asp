<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vTop, vSDate, vEDate, vGubun, vParam, vItemID, vMakerID
	vGubun		= NullFillWith(request("gubun"),"i")
	vSDate		= NullFillWith(Request("sdate"),"")
	vEDate		= NullFillWith(Request("edate"),"")
	vItemID		= NullFillWith(Request("itemid"),"")
	vMakerID	= NullFillWith(Request("makerid"),"")
	
	If vSDate = "" Then
		vSDate = Left(formatdatetime(DateAdd("yyyy",-1,now()),0),10)
	End IF
	
	If vEDate = "" Then
		vEDate = Left(formatdatetime(now(),0),10)
	End IF
	
	vParam = "gubun="&vGubun&"&sdate="&vSDate&"&edate="&vEDate&"&itemid="&vItemID&"&makerid="&vMakerID&""
	vParam = Replace(vParam, "&", "^^")
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript" src="/admin/datamart/fusionchart/FusionCharts.js"></script>

<script language="JavaScript">
function checkform(frm1)
{
	/*
	if(frm1.gubun.value != "m")
	{
		if(isNaN(frm1.itemid.value) && frm1.itemid.value != "")
		{
			alert("상품ID는 숫자로만 입력하세요.");
			frm1.itemid.value = "";
			return false;
		}
	}
	*/
}

</script>

<!-- 그래프 시작-->	

<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5 0 5 0"><center><font size="3">[<b>배송 소요일 분석</b>]</font></center></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<table width="850" align="center" cellpadding="0" cellspacing=0" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<form name="frm1" method="post" onSubmit="return checkform(this);">
		<tr bgcolor="#FFFFFF">
			<td>
				<table cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
				<tr bgcolor="#FFFFFF">
					<td width="760">
						&nbsp;
						기간:
				        <input id="sdate" name="sdate" value="<%=vSDate%>" class="text" size="10" maxlength="10" />
				        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
				        <input id="edate" name="edate" value="<%=vEDate%>" class="text" size="10" maxlength="10" />
				        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
						<script language="javascript">
							var CAL_Start = new Calendar({
								inputField : "sdate", trigger    : "sdate_trigger",
								onSelect: function() {
									var date = Calendar.intToDate(this.selection.get());
									CAL_End.args.min = date;
									CAL_End.redraw();
									this.hide();
								}, bottomBar: true, dateFormat: "%Y-%m-%d"
							});
							var CAL_End = new Calendar({
								inputField : "edate", trigger    : "edate_trigger",
								onSelect: function() {
									var date = Calendar.intToDate(this.selection.get());
									CAL_Start.args.max = date;
									CAL_Start.redraw();
									this.hide();
								}, bottomBar: true, dateFormat: "%Y-%m-%d"
							});
						</script>
					</td>
					<td width="90" valign="top" align="center"><input type="submit" value="Search"></td>
				</tr>
				<!--
				<tr id="searchitem" bgcolor="#FFFFFF" style="display:<% If vGubun = "i" OR vGubun = "io" Then Response.Write "block" Else Response.Write "none" End If %>;">
					<td>
						상품ID:
						<input type="text" class="text" name="itemid" value="<%=vItemID%>" size="7">
						<input type="button" class="button" value="찾기" onClick="popItemWindow('frm1')">
						&nbsp;&nbsp;
					</td>
				</tr>
				<tr id="searchmaker" bgcolor="#FFFFFF" style="display:<% If vGubun = "m" Then Response.Write "block" Else Response.Write "none" End If %>;">
					<td>
						브랜드:
					    <input type="text" class="text" name="makerid" value="<%=vMakerID%>" size="15" >
					    <input type="button" class="button" value="ID검색" onclick="jsSearchBrandID('frm1','makerid');" >
						&nbsp;&nbsp;
					</td>
				</tr>
				//-->
			    </table>
			</td>
		</tr>
		<tr height="10" bgcolor="#FFFFFF"><td></td></tr>
		</form>
		</table>
	</td>
</tr>
<tr>
	<td style="padding:5 0 5 0">
		<div id="chartdiv0" align="center"></div>
		<script type="text/javascript">	
		var chart = new FusionCharts("/admin/datamart/fusionchart/Area2D.swf", "chartdiv0", "850", "400", "0", "0");
		chart.setDataURL("/admin/datamart/fusionchart/graph_xml.asp?param=^^<%=vParam%>");
		chart.render("chartdiv0");
		</script>
	</td>
</tr>
</table>
<!-- 그래프 끝-->
