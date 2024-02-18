<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
Response.AddHeader "Content-Disposition","attachment;filename=�̺�Ʈ���_��_" & requestCheckVar(request("SType"),10) & "_" & date & hour(now) & minute(now) & ".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<%

dim SType '// �з�
dim EventID,ItemID, itemoption,i, makerid
dim BasicDateSet, Sdate, Edate, page
dim sortMethod

Dim oldlist


SType = requestCheckVar(request("SType"),10)
EventID = requestCheckVar(request("EventID"),10)
ItemID = requestCheckVar(request("ItemID"),10)
itemoption = requestCheckVar(request("itemoption"),10)  ''2013/10/14 �߰�
oldlist = requestCheckVar(request("oldlist"),10)
makerid = requestCheckVar(request("makerid"),32)

Sdate = requestCheckVar(request("Sdate"),10)
Edate = requestCheckVar(request("Edate"),10)

sortMethod = requestCheckVar(request("sortMethod"),8)
if sortMethod="" then sortMethod="totNoDS"

IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now())+1)
End IF

dim  oReport  '// ��� ����Ÿ
	set oReport = new CReportMaster
	oReport.FRectEventID = EventID
	oReport.FRectItemID = ItemID
	oReport.FRectMakerid = makerid
	oReport.FRectItemOption = ItemOption
	oReport.FRectStart = Sdate
	oReport.FRectEnd =  dateSerial(year(Edate),month(EDate),Day(EDate))
	oReport.FRectOldJumun = oldlist

dim t_TotalCost, t_FTotalNo
t_TotalCost = 0
t_FTotalNo  = 0
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<table width="1000" cellspacing="1" class="a" bgcolor="#DDDDFF">
<%

SELECT CASE SType

	CASE "D" '// ��¥�� �̺�Ʈ ���
	    IF (ItemID<>"") then
	        call oReport.GetEventStatisticsByDate
		ELSE
		    call oReport.GetEventStatisticsByDateDataMart
		END IF
%>
		<tr bgcolor="#DDDDFF">
	    	<td align="center">������</td>
	    	<td align="center">�Ǹž�</td>
			<td align="center">�ǸŰ���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fselldate %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>��</td>
   </tr>
		<% next %>
	<% end if %>

<%
	CASE "T"  '// ��ǰ�� �̺�Ʈ ���
		oReport.FRectSort = sortMethod
		call oReport.GetEventStatisticsByItemIDDataMart
%>
		<tr bgcolor="#EDEDFF">
			<td width="150" align="center" rowspan="2">�귣��</td>
			<td width="90" align="center" rowspan="2">�����۹�ȣ</td>
			<td width="70" align="center" colspan="2">��</td>
			<td width="70" align="center" colspan="2">PC��</td>
			<td width="70" align="center" colspan="2">�������</td>
			<td width="70" align="center" colspan="2">APP</td>
			<td width="70" align="center" colspan="2">���޸�</td>
			<td width="70" align="center" rowspan="2">Wish</td>
		</tr>
		<tr bgcolor="#EDEDFF">
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
			<td align="center"><%= oReport.FMasterItemList(i).FItemid %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_PC,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_mobile,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_App,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_App,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_outmall,0) %>��</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FwishCnt,0) %>��</td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "O"  '// �ɼǺ� �̺�Ʈ ���
		call oReport.GetEventStatisticsByItemOptionDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td align="center">�����۹�ȣ</td>
			<td align="center">�ɼǹ�ȣ</td>
			<td align="center">�Ǹž�</td>
			<td align="center">�ǸŰ���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).FItemid %></td>
			<td align="center"><%= oReport.FMasterItemList(i).FItemOption %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>��</td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "M"  '// �귣�庰 �̺�Ʈ ���
		call oReport.GetEventStatisticsByMakerIDDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td align="center">�귣��</td>
			<td align="center">�Ǹž�</td>
			<td align="center">�ǸŰ���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>��</td>
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
			<td align="right"><%= FormatNumber(t_TotalCost,0) %></td>
			<td align="right"><%= FormatNumber(t_FTotalNo,0) %> ��</td>
		</tr>
	</table>

<%
	set oReport = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
