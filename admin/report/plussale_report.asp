<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/plussale_reportcls.asp"-->
<%

dim SType '// �з� 
dim i, cateNo
dim BasicDateSet, Sdate, Edate, page, grpWidth
dim pTT, pCT, nTT, nCT

SType = request("SType")

Sdate = request("Sdate")
Edate = request("Edate")
cateNo = requestCheckVar(request("cateNo"),10)

if SType="" then SType="T"	'��ǰ�� �⺻
IF Sdate="" THEN
	Sdate= dateadd("ww",-1,date())
End IF

IF Edate="" THEN
	Edate= date()
End IF
	
dim  oReport  '// ��� ����Ÿ 
	set oReport = new CReportMaster
	oReport.FRectCateNo = cateNo
	oReport.FRectStart = Sdate
	oReport.FRectEnd =  Edate
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

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		�˻� �Ⱓ : 
			<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
			<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
		<br>
		ī�װ����� : <% DrawSelectBoxCategoryLarge "cateNo",cateNo %>&nbsp;
		�з� : 
			<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>>��¥��
			<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>>��ǰ��
		</td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
	
<table width="100%" cellspacing="1" class="a" bgcolor="#DDDDFF">

<%

SELECT CASE SType
	
	CASE "D" '// ��¥�� ������� 
		call oReport.GetSaleStatisticsByDate
%>
		<tr bgcolor="#DDDDFF">
	    	<td align="center">������</td>
			<td align="center">�÷��������</td>
			<td align="center">�÷����Ǹż�</td>
			<td align="center">�Ѹ����</td>
			<td align="center">���Ǹż�</td>
			<td align="center">�÷�������</td>
		</tr>
	<%
		if oReport.FResultCount > 0 then
			for i=0 to oReport.FResultCount-1
				'���� ���
				pTT = pTT + oReport.FMasterItemList(i).FsellPlustotal
				pCT = pCT + oReport.FMasterItemList(i).FsellPluscnt
				nTT = nTT + oReport.FMasterItemList(i).Fselltotal
				nCT = nCT + oReport.FMasterItemList(i).Fsellcnt
	%>
		<tr bgcolor="#FFFFFF">
		<td align="center"><%= oReport.FMasterItemList(i).Fselldate %></td>
		<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FsellPlustotal,0) %></td>
		<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FsellPluscnt,0) %>��</td>
		<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
		<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %>��</td>
		<td align="center">
	        <% if oReport.FMasterItemList(i).Fselltotal<>0 then %>
	        <%= FormatPercent(oReport.FMasterItemList(i).FsellPlustotal/oReport.FMasterItemList(i).Fselltotal) %>
	        <% end if %>
		</td>
	   </tr>
		<% next %>
		<tr>
			<td align="center"><b>����</b></td>
			<td align="right"><b><%= FormatNumber(pTT,0) %></b></td>
			<td align="right"><b><%= FormatNumber(pCT,0) %> ��</b></td>
			<td align="right"><b><%= FormatNumber(nTT,0) %></b></td>
			<td align="right"><b><%= FormatNumber(nCT,0) %> ��</b></td>
			<td align="center"><b><% if nTT<>0 then Response.Write FormatPercent(pTT/nTT) %></b></td>
		</tr>
	<% end if %>
<% 
	CASE "T"  '// ��ǰ�� ������� 
		call oReport.GetSaleStatisticsByItemID
%>
		<tr bgcolor="#DDDDFF">
			<td align="center">��ǰ�ڵ�</td>
			<td align="center">�̹���</td>
			<td align="center">�귣��ID</td>
			<td align="center">��ǰ��</td>
			<td align="center">�÷��������</td>
			<td align="center">�÷����Ǹż�</td>
			<td align="center">�Ѹ����</td>
			<td align="center">���Ǹż�</td>
			<td align="center">�÷�������</td>
		</tr>
	<%
		if oReport.FResultCount > 0 then
			for i=0 to oReport.FResultCount-1
				'���� ���
				pTT = pTT + oReport.FMasterItemList(i).FsellPlustotal
				pCT = pCT + oReport.FMasterItemList(i).FsellPluscnt
				nTT = nTT + oReport.FMasterItemList(i).Fselltotal
				nCT = nCT + oReport.FMasterItemList(i).Fsellcnt
	%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).FItemid %></td>
			<td align="center"><img src="<%= oReport.FMasterItemList(i).FSmallImage %>" width="50"></td>
			<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
			<td align="left"><%= oReport.FMasterItemList(i).FitemName %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FsellPlustotal,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FsellPluscnt,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %>��</td>
			<td align="center">
		        <% if oReport.FMasterItemList(i).Fselltotal<>0 then %>
		        <%= FormatPercent(oReport.FMasterItemList(i).FsellPlustotal/oReport.FMasterItemList(i).Fselltotal) %>
		        <% end if %>
			</td>
		</tr>
		<% next %>
		<tr>
			<td colspan="4" align="center"><b>����</b></td>
			<td align="right"><b><%= FormatNumber(pTT,0) %></b></td>
			<td align="right"><b><%= FormatNumber(pCT,0) %> ��</b></td>
			<td align="right"><b><%= FormatNumber(nTT,0) %></b></td>
			<td align="right"><b><%= FormatNumber(nCT,0) %> ��</b></td>
			<td align="center"><b><% if nTT<>0 then Response.Write FormatPercent(pTT/nTT) %></b></td>
		</tr>
	<% end if %>
<%
	CASE ELSE
		response.write "�����߻�,�ٽ� �õ�"
END SELECT
%>
	</table>	

<%
set oReport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->