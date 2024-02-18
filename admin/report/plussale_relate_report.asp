<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/plussale_reportcls.asp"-->
<%

dim itemid,i,cateNo

dim BasicDateSet, Sdate, Edate, page
Dim ttSellPrice

itemid = requestCheckVar(request("itemid"),8)

Sdate = requestCheckVar(request("Sdate"),10)
Edate = requestCheckVar(request("Edate"),10)

cateNo = requestCheckVar(request("cateNo"),10)

IF Sdate="" THEN
	Sdate= dateadd("ww",-1,date())
End IF

IF Edate="" THEN
	Edate= date()
End IF


dim oReport
set oReport = new CReportMaster
oReport.FRectStart = Sdate
oReport.FRectEnd =  Edate
oReport.FRectCateNo = cateNo
oReport.FRectItemId = itemid
oReport.GetSaleStatisticsAll
%>
<script language="javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	function changecontent(){
		document.frm.submit();
	}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#DDDDFF">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		�˻� �Ⱓ : 
			<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
			<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
		ī�װ����� : <% DrawSelectBoxCategoryLarge "cateNo",cateNo %>&nbsp;
		��ǰ��ȣ : <input type="text" size="10" name="itemid" value="">
		</td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
<br>

<% if oReport.FResultCount > 0 then %>
<table width="100%" cellspacing="0" class="a">
<tr>
	<td>�� ��ǰ�� : <%=oReport.FResultCount%>��</td>
	<td align="right">
		�Ѹ���� :
		<%
			ttSellPrice = 0
			for i=0 to oReport.FResultCount-1
				ttSellPrice = ttSellPrice + oReport.FMasterItemList(i).Fselltotal
			next
			Response.Write FormatNumber(ttSellPrice,0)
		%>�� /
		����ո���� : <%=FormatNumber(ttSellPrice/oReport.FResultCount,0) %>��
	</td>
</tr>
</table>
<table width="100%" cellspacing="1" class="a" bgcolor="#3d3d3d">
	<tr bgcolor="#DDDDFF">
		<td align="center" width="40">��ǰ��ȣ</td>
		<td align="center" width="50">�̹���</td>
		<td align="center" width="100">�귣��ID</td>
		<td align="center" >��ǰ��</td>
		<td align="center" width="80">�� ����</td>
		<td align="center" width="80">�� �Ǹż�</td>
		<td align="center" width="100">�� ���� </td>
	</tr>
	<% for i=0 to oReport.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= oReport.FMasterItemList(i).FItemid %></td>
		<td align="center"><img src="<%= oReport.FMasterItemList(i).FSmallImage %>" width="50"></td>
		<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
		<td align="left"><%= oReport.FMasterItemList(i).FitemName %></td>
		<td align="center"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
		<td align="center"><%= oReport.FMasterItemList(i).Fsellcnt %></td>
		<td align="center">
			<a href="/admin/report/plussale_report_detail.asp?SType=D&itemid=<%= oReport.FMasterItemList(i).FItemid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">��¥��</a>
			 | 
			<a href="/admin/report/plussale_report_detail.asp?SType=T&itemid=<%= oReport.FMasterItemList(i).FItemid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">��ǰ��</a>
		</td>
	</tr>
	<% next %>
</table>
<% else %>
<table width="800" cellspacing="1" class="a" bgcolor="#3d3d3d">
	<tr bgcolor="#DDDDFF">
		<td align="center"> [ ����� �����ϴ�]
		</td>
	</tr>
	
</table>
<% end if %>
<%
set oReport = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->