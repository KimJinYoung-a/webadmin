<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim yyyy1 : yyyy1=RequestCheckVar(request("yyyy1"),4)
Dim yyyy2 : yyyy2=RequestCheckVar(request("yyyy2"),4)
Dim mm1 : mm1=RequestCheckVar(request("mm1"),2)
Dim mm2 : mm2=RequestCheckVar(request("mm2"),2)

Dim ckMinus : ckMinus=RequestCheckVar(request("ckMinus"),10)

Dim cdl : cdl=RequestCheckVar(request("cdl"),3)
Dim cdm : cdm=RequestCheckVar(request("cdm"),3)
Dim cds : cds=RequestCheckVar(request("cds"),3)
Dim catebase : catebase=RequestCheckVar(request("catebase"),10)

Dim chartXMLURL : chartXMLURL="/admin/datamart/fusionchart/chartResponse.asp?mode=C1&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&ckMinus="&ckMinus&"&yyyy1="&yyyy1&"&mm1="&mm1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&catebase="&catebase

''rw chartXMLURL
%>

<script language="javascript" src="/admin/maechul/fusionchart/FusionCharts.js"></script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="cd1" value="">
	<input type="hidden" name="cd2" value="">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %>
		&nbsp;&nbsp;
		<select name="ckMinus">
		<option value="" >반품포함
		<option value="1" <%= CHKIIF(ckMinus="1","selected","") %> >반품제외
		<option value="2" <%= CHKIIF(ckMinus="2","selected","") %> >반품주문만
		</select>
		&nbsp;
		카테고리:
		<% SelectBoxCategoryLarge cdl %>

		&nbsp;
		카테고리 매출 기준:
		<input type="radio" name="catebase" value="S" <%= CHKIIF(catebase="S","checked","") %> >판매시카테고리
		<input type="radio" name="catebase" value="C" <%= CHKIIF(catebase="C","checked","") %> >현재카테고리
		<td class="a" align="right">
			<a href="javascript:frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>


<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5 0 5 0"><center><font size="3"></font></center></td>
</tr>
<tr>
	<td style="padding:5 0 5 0">
		<div id="chartdiv0" align="center"></div>
		<script type="text/javascript">
		var chart = new FusionCharts("/admin/maechul/fusionchart/MSLine.swf", "chartdiv0", "900", "500", "0", "0");
		chart.setDataURL("<%= Server.URLEnCode(chartXMLURL) %>");
		chart.render("chartdiv0");
		</script>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
