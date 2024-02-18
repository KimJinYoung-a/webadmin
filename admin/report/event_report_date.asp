<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
''response.write "<H3>������...</H3>"
'dbget.close()	:	response.End

dim eventid,i,sKind,cateNo,ReportType
Dim dispCate
dim BasicDateSet, Sdate, Edate, page
Dim oldlist, ttSellPrice
dim strSort
 

ReportType = requestCheckVar(request("rt"),10)
eventid = requestCheckVar(request("eventid"),6)

Sdate = requestCheckVar(request("Sdate"),10)
Edate = requestCheckVar(request("Edate"),10)

oldlist = requestCheckVar(request("oldlist"),10)

cateNo = requestCheckVar(request("cateNo"),10)
sKind = requestCheckVar(Request("eventkind"),10)	'�̺�Ʈ����
dispCate	= requestCheckVar(Request("disp"),10) 		'���� ī�װ�

strSort = requestCheckVar(Request("selSort"),3) 
if strSort = "" then strSort ="TMD"
IF ReportType="" THEN ReportType="s"

IF sKind = "" THEN
	sKind="1"
END IF


IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now()))
End IF
Function DateToWeekName(d)
	SELECT CASE d
		CASE "1" : DateToWeekName = "<font color=""red"">��</font>"
		CASE "2" : DateToWeekName = "��"
		CASE "3" : DateToWeekName = "ȭ"
		CASE "4" : DateToWeekName = "��"
		CASE "5" : DateToWeekName = "��"
		CASE "6" : DateToWeekName = "��"
		CASE "7" : DateToWeekName = "<font color=""blue"">��</font>"
	END SELECT
End Function

dim oReport
set oReport = new CReportMaster
oReport.FRectStart = Sdate
oReport.FRectEnd =  dateSerial(year(Edate),month(EDate),Day(EDate)+1)
oReport.FRectOldJumun = oldlist
oReport.FRectCateNo = cateNo
oReport.FRectDispCate = dispCate
oReport.FRectEventid = eventid
oReport.FRectEvtKind = sKind
oReport.FRectReportType= ReportType
oReport.FRectSort = strSort

'// 2014-08-27, skyer9
'if (DateDiff("m", Sdate, dateSerial(year(Edate),month(EDate),Day(EDate)))) > 1 then
'	response.write "�ѹ��� 2�� �̻��� �˻��� �� �����ϴ�."
'	dbget.close()
'	response.end
'end if

oReport.GetEventStatisticsByDateDataMart_New

 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	function changecontent(){
		document.frm.submit();
	}
	
	//����Ʈ ����
function jstrSort(sValue,i){
	 	document.frm.selSort.value= sValue;

		   if (-1 < eval("document.frmList.img"+i).src.indexOf("_alpha")){
	        document.frm.selSort.value= sValue+"D";
	    }else if (-1 < eval("document.frmList.img"+i).src.indexOf("_bot")){
	     		document.frm.selSort.value= sValue+"A";
	    }else{
	       document.frm.selSort.value= sValue+"D";
	    }
		 document.frm.submit();
	}
</script>


	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="selSort" value="<%=strSort%>"><!--����-->
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
		<td align="left">
			<table class="a" border="0" cellpadding="3">
			<tr>
			<td class="a" >
				 
			 �Ⱓ:
					<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
					<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');"> 
			</td>
		</tr>
		<tr>
			<td>
				�̺�Ʈ ���� <%sbGetOptEventCodeValue "eventkind", sKind, True,""%>
				&nbsp;�̺�Ʈ�ڵ� : <input type="text" size="10" name="eventid" value="<%=eventid%>">
				&nbsp;����ī�װ�: <% DrawSelectBoxCategoryLarge "cateNo",cateNo %>
				&nbsp;����ī�װ�: <%=fnDispCateSelectBox(1,"","disp",dispCate,"") %>
			</td>
		</tr>
		</table>
	</td>
		<td class="a" align="center" bgcolor="<%= adminColor("gray") %>"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</table>
	 </form> 
<br>
<% if oReport.FResultCount > 0 then %>
<table width="100%" cellspacing="0" cellpadding="3" class="a">    
<tr>
	<td>���̺�Ʈ �� : <%=oReport.FResultCount%>��</td>
	<td align="right">
		���̺�Ʈ����� :
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
<form name="frmList" method="post">
<table width="100%" cellspacing="1" cellpadding="5" class="a" bgcolor="#3d3d3d">
	<tr bgcolor="#DDDDFF" align="center">
		<td width="40" rowspan="2" colspan="2"> �Ⱓ</td>  
		<td colspan="4">Mobile/App</td>
		<td colspan="4"> PC-Web </td>
		<td colspan="4">����</td>
		<td colspan="4">3PL</td>
	    <td rowspan="2" >�� �Ǹż�</td>
		<td rowspan="2">�����հ�</td>
		<td rowspan="2" >����</td>  
	</tr>
	<tr bgcolor="#DDDDFF" align="center">
	    <td>�Ǹż�</td>
	    <td>����</td>
	    <td>������</td>
	    <td>����</td>  
	    <td>�Ǹż�</td>
	    <td>����</td> 
	    <td>������</td>
	    <td>����</td> 
	    <td>�Ǹż�</td>
	    <td>����</td>
	    <td>������</td>
	    <td>����</td>  
	    <td>�Ǹż�</td>
	    <td>����</td>
	    <td>������</td>
	    <td>����</td>  
	</tr>
	<tr bgcolor="#EEEEEE"  align="right">
	    <td colspan="2" align="center">���հ�</td>
	    <td><%= FormatNumber(oReport.FTotCnt_m,0) %></td>
		<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FTotSell_m,0) %></b></td>
		<td bgcolor="#DDFFDD"><b><%if oReport.FTotSell_m > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_m/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FTotSell_m -oReport.FTotBuy_m,0) %></td> 
	 
	    <td><%= FormatNumber(oReport.FTotCnt_p,0) %></td>
	    <td bgcolor="#EEEEEE"><b><%= FormatNumber(oReport.FTotSell_p,0) %></b></td>
		<td bgcolor="#EEEEEE"><b><%if oReport.FTotSell_p > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_p/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FTotSell_p -oReport.FTotBuy_p,0) %></td> 
		
		<td><%= FormatNumber(oReport.FTotCnt_o,0) %></td>
		<td ><b><%= FormatNumber(oReport.FTotSell_o,0) %></b></td>
		<td ><b><%if oReport.FTotSell_o > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_o/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FTotSell_o -oReport.FTotBuy_o,0) %></td> 
		
		<td><%= FormatNumber(oReport.FTotCnt_3,0) %></td>
		<td ><b><%= FormatNumber(oReport.FTotSell_3,0) %></b></td>
		<td ><b><%if oReport.FTotSell_3 > 0 and oReport.FTotSell > 0 then %><%= FormatNumber((oReport.FTotSell_3/oReport.FTotSell)*100,0) %>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FTotSell_3 -oReport.FTotBuy_3,0) %></td> 
		
		<td><%= FormatNumber(oReport.FTotCnt,0) %></td>
		<td><b><%= FormatNumber(oReport.FTotSell,0) %></b></td>
		<td><b><%=FormatNumber(oReport.FTotSell-oReport.FTotBuy,0)%></b></td>
	         
        
	</tr> 
	<% for i=0 to oReport.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" align="right">
		<td align="center">
		    <%IF not isNull(oReport.FMasterItemList(i).FYYYYMMDD) then %>
			<% if right(FormatDateTime(oReport.FMasterItemList(i).FYYYYMMDD,1),3) = "�����" then %>
				<font color="blue"><%=oReport.FMasterItemList(i).FYYYYMMDD %></font>
			<% elseif right(FormatDateTime(oReport.FMasterItemList(i).FYYYYMMDD,1),3) = "�Ͽ���" then %>
				<font color="red"><%= oReport.FMasterItemList(i).FYYYYMMDD %></font>
			<% else %>
				<%= oReport.FMasterItemList(i).FYYYYMMDD %>
			<% end if %>
			<% end if %>
		</td>
		<td align="center">  <%IF not isNull(oReport.FMasterItemList(i).FYYYYMMDD) then %><%= DateToWeekName(DatePart("w",oReport.FMasterItemList(i).FYYYYMMDD)) %><% end if %></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_mobile,0) %></td>
		<td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile,0) %></b></td>
		<td bgcolor="#DDFFDD"><b><%if oReport.FMasterItemList(i).Fsellsum_mobile > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber((oReport.FMasterItemList(i).Fsellsum_mobile/oReport.FMasterItemList(i).Fselltotal)*100,0) %>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile -oReport.FMasterItemList(i).Fbuysum_mobile,0) %></td> 
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_PC,0) %></td>
		<td bgcolor="#EEEEEE"><b><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC,0) %></b></td>
		<td bgcolor="#EEEEEE"><b><%if  oReport.FMasterItemList(i).Fsellsum_PC > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %> <%=FormatNumber((oReport.FMasterItemList(i).Fsellsum_PC/oReport.FMasterItemList(i).Fselltotal)*100,0)%>%<%end if%></b></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC -oReport.FMasterItemList(i).Fbuysum_PC,0) %></td> 
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_outmall,0) %></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall,0) %></td>
		<td><%if oReport.FMasterItemList(i).Fsellsum_outmall > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall/oReport.FMasterItemList(i).Fselltotal*100,0) %>%<%end if%></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall -oReport.FMasterItemList(i).Fbuysum_outmall,0) %></td> 
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_3PL,0) %></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL,0) %></td>
		<td><%if oReport.FMasterItemList(i).Fsellsum_3PL > 0 and oReport.FMasterItemList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL/oReport.FMasterItemList(i).Fselltotal*100,0) %>%<%end if%></td>
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_3PL -oReport.FMasterItemList(i).Fbuysum_3PL,0) %></td> 
		<td><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %></td>
		<td><b><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></b></td>
		<td><b><%=FormatNumber(oReport.FMasterItemList(i).Fselltotal-oReport.FMasterItemList(i).Fbuytotal,0)%></b></td>  
	</tr>
	<% next %>
</table>
</form>
<% else %>
<table width="100%" cellspacing="1"  cellpadding="3" class="a" bgcolor="#3d3d3d">
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
