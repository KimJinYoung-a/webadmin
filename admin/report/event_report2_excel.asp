<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
Response.AddHeader "Content-Disposition","attachment;filename=�̺�Ʈ���_" & date & hour(now) & minute(now) & ".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<%
'###########################################################
' Description : �̺�Ʈ���
' Hieditor : ������ ����
'			 2021.02.23 �ѿ�� ����(�˻����� �߰�. �ֱٵ��������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim eventid,i,sKind,cateNo,ReportType, dispCate, reloading, oReport
dim BasicDateSet, Sdate, Edate, page, oldlist, ttSellPrice, strSort, eType
	ReportType = requestCheckVar(request("rt"),10)
	eventid = requestCheckVar(request("eventid"),6)
	Sdate = requestCheckVar(request("Sdate"),10)
	Edate = requestCheckVar(request("Edate"),10)
	oldlist = requestCheckVar(request("oldlist"),10)
	cateNo = requestCheckVar(request("cateNo"),10)
	sKind = requestCheckVar(Request("eventkind"),10)	'�̺�Ʈ����
	eType = requestCheckVar(Request("eventtype"),10)	'�̺�Ʈ����
	dispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�
	strSort = requestCheckVar(Request("selSort"),3)
    reloading = requestCheckVar(request("reloading"),2)

if strSort = "" then strSort ="TMD"
IF ReportType="" THEN ReportType="s"

IF reloading="" and sKind = "" THEN sKind="1"
IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF
IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

set oReport = new CReportMaster
	oReport.FRectStart = Sdate
	oReport.FRectEnd =  dateSerial(year(Edate),month(EDate),Day(EDate)+1)
	oReport.FRectOldJumun = oldlist
	oReport.FRectCateNo = cateNo
	oReport.FRectDispCate = dispCate
	oReport.FRectEventid = eventid
	oReport.FRectEvtKind = sKind
	oReport.FRectEvtType = eType
	oReport.FRectReportType= ReportType
	oReport.FRectSort = strSort

	'// 2014-08-27, skyer9
	if (DateDiff("m", Sdate, dateSerial(year(Edate),month(EDate),Day(EDate)))) > 1 then
		response.write "�ѹ��� 2�� �̻��� �˻��� �� �����ϴ�."
		dbget.close()
		response.end
	end if

	oReport.GetEventStatisticsDataMart

%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		�˻���� : <b><%= oReport.FResultCount %></b>
		<% if oReport.FResultCount > 0 then %>
			&nbsp;
			���̺�Ʈ����� :
			<%
				ttSellPrice = 0
				for i=0 to oReport.FResultCount-1
					ttSellPrice = ttSellPrice + oReport.FMasterItemList(i).Fselltotal
				next
				Response.Write FormatNumber(ttSellPrice,0)
			%>�� /
			����ո���� : <%=FormatNumber(ttSellPrice/oReport.FResultCount,0) %>��
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60" rowspan="2"><b>�̺�Ʈ��ȣ</b></td> 
	<td rowspan="2">�̺�Ʈ��</td>
	<td  colspan="5">Mobile/App</td>
	<td colspan="5"> PC-Web </td>
	<td colspan="5">����</td>
	<td colspan="5">3PL</td>
	<td  rowspan="2" >�� �Ǹż�</td>
	<td  rowspan="2"><b>�����հ�</b></td>
	<td  rowspan="2"><b>��޾�</b></td>
	<td  rowspan="2"><b>����</b></td> 
	<td width="160" rowspan="2">�̺�Ʈ �Ⱓ</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Ǹż�</td>
	<td><b>����</b></td>
	<td><b>��޾�</b></td>
	<td>������</td>
	<td><b>����</b></td>  
	<td>�Ǹż�</td>
	<td><b>����</b></td> 
	<td><b>��޾�</b></td>
	<td>������</td>
	<td><b>����</b></td> 
	<td>�Ǹż�</td>
	<td><b>����</b></td>
	<td><b>��޾�</b></td>
	<td>������</td>
	<td><b>����</b></td> 
	<td>�Ǹż�</td>
	<td><b>����</b></td>
	<td><b>��޾�</b></td>
	<td>������</td> 
	<td><b>����</b></td>  
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="2" align="center">���հ�</td>
	<td><%= FormatNumber(oReport.FTotCnt_m,0) %></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FTotSell_m,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FTotreducedprice_m,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%if oReport.FTotSell_m > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_m/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FTotSell_m -oReport.FTotBuy_m,0) %></td> 
	
	<td><%= FormatNumber(oReport.FTotCnt_p,0) %></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FTotSell_p,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FTotreducedprice_p,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%if oReport.FTotSell_p > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_p/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FTotSell_p -oReport.FTotBuy_p,0) %></td> 
	
	<td><%= FormatNumber(oReport.FTotCnt_o,0) %></td>
	<td ><b><%= FormatNumber(oReport.FTotSell_o,0) %></b></td>
	<td ><b><%= FormatNumber(oReport.FTotreducedprice_o,0) %></b></td>
	<td ><b><%if oReport.FTotSell_o > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_o/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FTotSell_o -oReport.FTotBuy_o,0) %></td> 
	
	<td><%= FormatNumber(oReport.FTotCnt_3,0) %></td>
	<td ><b><%= FormatNumber(oReport.FTotSell_3,0) %></b></td>
	<td ><b><%= FormatNumber(oReport.FTotreducedprice_3,0) %></b></td>
	<td ><b><%if oReport.FTotSell_3 > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_3/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FTotSell_3 -oReport.FTotBuy_3,0) %></td> 
	
	<td><%= FormatNumber(oReport.FTotCnt,0) %></td>
	<td><b><%= FormatNumber(oReport.FTotSell,0) %></b></td>
	<td><b><%= FormatNumber(oReport.FTotreducedprice,0) %></b></td>
	<td><b><%=FormatNumber(oReport.FTotSell-oReport.FTotBuy,0)%></b></td>
	<td></td>
</tr>
<% if oReport.FResultCount > 0 then %>
<% for i=0 to oReport.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="center"><a href="<%= wwwURL %>/event/eventmain.asp?eventid=<%= oReport.FMasterItemList(i).FEventIdx %>" target="_blank"><%= oReport.FMasterItemList(i).FEventIdx %></a></td>
	<td align="left"><a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oReport.FMasterItemList(i).FEventIdx %>" target="_blank"><%= oReport.FMasterItemList(i).FEventName %></a></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_mobile,0) %></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FMasterItemList(i).freducedprice_Mobile,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%if oReport.FMasterItemList(i).Fsellsum_mobile > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber((oReport.FMasterItemList(i).Fsellsum_mobile/oReport.FMasterItemList(i).Fselltotal)*100,0) %>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile -oReport.FMasterItemList(i).Fbuysum_mobile,0) %></td> 
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_PC,0) %></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FMasterItemList(i).freducedprice_PC,0) %></b></td>
	<td bgcolor="#DDFFDD"><b><%if  oReport.FMasterItemList(i).Fsellsum_PC > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %> <%=FormatNumber((oReport.FMasterItemList(i).Fsellsum_PC/oReport.FMasterItemList(i).Fselltotal)*100,0)%>%<%end if%></b></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC -oReport.FMasterItemList(i).Fbuysum_PC,0) %></td> 
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_outmall,0) %></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall,0) %></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).freducedprice_Outmall,0) %></td>
	<td><%if oReport.FMasterItemList(i).Fsellsum_outmall > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall/oReport.FMasterItemList(i).Fselltotal*100,0) %>%<%end if%></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall -oReport.FMasterItemList(i).Fbuysum_outmall,0) %></td> 
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_3PL,0) %></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL,0) %></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).freducedprice_3PL,0) %></td>
	<td><%if oReport.FMasterItemList(i).Fsellsum_3PL > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL/oReport.FMasterItemList(i).Fselltotal*100,0) %>%<%end if%></td>
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL -oReport.FMasterItemList(i).Fbuysum_3PL,0) %></td> 
	<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %></td>
	<td><b><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></b></td>
	<td><b><%= FormatNumber(oReport.FMasterItemList(i).fTotalreducedprice,0) %></b></td>
	<td><b><%=FormatNumber(oReport.FMasterItemList(i).Fselltotal-oReport.FMasterItemList(i).Fbuytotal,0)%></b></td>
	
	<td align="center">
		<%= oReport.FMasterItemList(i).FStartDay %> ~ <%= oReport.FMasterItemList(i).FEndDay %>
	</td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="30" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oReport = Nothing

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
