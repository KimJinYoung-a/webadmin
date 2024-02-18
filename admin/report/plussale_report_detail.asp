<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/plussale_reportcls.asp"-->
<%

dim SType '// �з� 
dim ItemID, i
dim BasicDateSet, Sdate, Edate, page, grpWidth

SType = request("SType")
ItemID = request("ItemID")

Sdate = request("Sdate")
Edate = request("Edate")

IF Sdate="" THEN
	Sdate= dateadd("ww",-1,date())
End IF

IF Edate="" THEN
	Edate= date()
End IF



	
dim  oReport  '// ��� ����Ÿ 
	set oReport = new CReportMaster
	oReport.FRectItemID = ItemID
	oReport.FRectStart = Sdate
	oReport.FRectEnd =  Edate

dim oTotal '// ���հ� 
	set oTotal = new CReportMaster
	oTotal.FRectItemID = ItemID
	oTotal.FRectStart = Sdate
	oTotal.FRectEnd =  Edate
	oTotal.GetSaleStatisticsTotal
	
	
%>

<script language="javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
</script>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		�˻� �Ⱓ : 
			<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
			<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
		<br>
		
		��ǰ ��ȣ : 
			<input type="text" name="ItemID" size="10" value="<%= ItemID %>">

		�з� : 
			<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>>��¥��
			<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>>��ǰ��
		</td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
	
<table width="800" cellspacing="1" class="a" bgcolor="#DDDDFF">

<%

SELECT CASE SType
	
	CASE "D" '// ��¥�� ������� 
		call oReport.GetSaleStatisticsByDate
%>
		<tr bgcolor="#DDDDFF">
	    	<td width="90" align="center">������</td>
	    	<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="500" align="center">�׷���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fselldate %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>��</td>
			<td width="500">
				<%
					'�׷��� ���� ��� (2008.07.08;������ ����)
					if oReport.maxc>0 then
						grpWidth = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxc*400)
					else
						grpWidth = 0
					end if
				%>
				<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
			</td>
	   </tr>
		<% next %>
	<% end if %>
	
<% 
	CASE "T"  '// ��ǰ�� ������� 
		call oReport.GetSaleStatisticsByItemID
%>
		<tr bgcolor="#DDDDFF">
			<td width="90" align="center">�����۹�ȣ</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="500" align="center">�׷���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).FItemid %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>��</td>
			<td>
				<%
					'�׷��� ���� ��� (2008.07.08;������ ����)
					if oReport.maxc>0 then
						grpWidth = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxc*400)
					else
						grpWidth = 0
					end if
				%>
				<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
			</td>
		</tr>
		<% next %>
	<% end if %>

<%
	CASE ELSE
		response.write "�����߻�,�ٽ� �õ�"
END SELECT
%>
		<tr>
			<td align="center">����</td>
			<td align="right"><%= FormatNumber(oTotal.FTotalCost,0) %></td>
			<td align="right"><%= FormatNumber(oTotal.FTotalNo,0) %> ��</td>
		</tr>
	</table>	

<%
set oReport = Nothing
set oTotal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->