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
Dim catebase : catebase=RequestCheckVar(request("catebase"),10)

Dim dispCate : dispCate = requestCheckvar(request("disp"),16)
Dim cdl : cdl=Left(dispCate,3)
Dim cdm : cdm=Mid(dispCate,4,3)
Dim cds : cds=Mid(dispCate,7,3)
Dim cdx : cdx=Mid(dispCate,10,3)

Dim chartXMLURL : chartXMLURL="/admin/datamart/fusionchart/chartResponse.asp?mode=C1&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&cdx="&cdx&"&ckMinus="&ckMinus&"&yyyy1="&yyyy1&"&mm1="&mm1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&catebase="&catebase

''rw chartXMLURL
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" src="/admin/maechul/fusionchart/FusionCharts.js"></script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="catebase" value="<%=catebase%>">
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
		전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->

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
