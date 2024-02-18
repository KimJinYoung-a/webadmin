<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<%

dim SType '// �з�
dim EventID,ItemID, itemoption,i, makerid
dim BasicDateSet, Sdate, Edate, page, grpWidth
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

'yyyy1 = request("yyyy1")
'mm1 = request("mm1")
'dd1 = request("dd1")

'yyyy2 = request("yyyy2")
'mm2 = request("mm2")
'dd2 = request("dd2")


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

'dim oTotal '// ���հ� ?? �ʿ�?
'	set oTotal = new CReportMaster
'	oTotal.FRectEventID = EventID
'	oTotal.FRectItemID = ItemID
'	oTotal.FRectStart = Sdate
'	oTotal.FRectEnd =  dateSerial(year(Edate),month(EDate),Day(EDate)+1)
'	oTotal.FRectOldJumun = oldlist

'	IF (ItemID<>"") then
'	    oTotal.GetEventStatisticsTotal
'	ELSe
'	    oTotal.GetEventStatisticsTotalDataMart
'	END IF

%>

<script type="text/javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function viewImage(div,itemid) {
		iframeDB1.location.href = "/admin/report/iframe_viewImage.asp?div="+div+"&itemid="+itemid+"";
	}

	function chgSortMethod(sm) {
		document.frm.target="_self";
		document.frm.action="";
		document.frm.sortMethod.value=sm;
		document.frm.submit();
	}

	function jsSubmit() {
		document.frm.target="_self";
		document.frm.action="";
		document.frm.submit();
	}

	// �����ޱ�
	function fnGetExcelFile() {
		document.frm.target="_blank";
		document.frm.action="/admin/report/event_report_detail_excel.asp";
		document.frm.submit();
	}
</script>

<table width="1000" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<input type="hidden" name="sortMethod" value="<%=sortMethod%>">
	<tr>
		<td class="a" >
		<!--
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6������������
		-->
		�˻� �Ⱓ :
			<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
			<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
		<br />

		�̺�Ʈ ��ȣ :
			<input type="text" name="EventID" size="10" value="<%= EventID %>">
        �귣�� :
			<input type="text" name="makerid" size="10" value="<%= makerid %>">
        ��ǰ ��ȣ :
            <input type="text" name="ITEMID" size="9" value="<%= ITEMID %>">
        �ɼ� ��ȣ :
            <input type="text" name="itemoption" size="9" value="<%= itemoption %>">
		<br />
		�з� :
			<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>> ��¥��
			<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>> ��ǰ��
			<input type="radio" name="SType" value="O" <% If SType = "O" Then response.write "checked" %>> �ɼǺ�
			<input type="radio" name="SType" value="M" <% If SType = "M" Then response.write "checked" %>> �귣�庰
		</td>
		<td class="a" align="right"><img src="/admin/images/search2.gif" width="74" height="22" border="0" onclick="jsSubmit();" style="cursor:pointer;"></td>
	</tr>
	<tr>
		<td colspan="2" style="background-color:#F4F4F4; text-align:right;">
			<img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" onclick="fnGetExcelFile()" style="cursor:pointer" />
		</td>
	<tr>
	</form>
</table>
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
	    	<td width="90" align="center">������</td>
	    	<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="500" align="center">�׷���</td>
			<td width="70" align="center">�󼼺���</td>
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
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=T&EventID=<%= EventID %>&SDate=<%=oReport.FMasterItemList(i).Fselldate%>&EDate=<%= oReport.FMasterItemList(i).Fselldate %>">����</a></td>
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
			<td width="90" align="center" rowspan="2" onClick="chgSortMethod('<%=chkIIF(SortMethod="itemidDS","itemidAS","itemidDS")%>')" style="cursor:pointer;">�����۹�ȣ<%=chkIIF(SortMethod="itemidDS","��",chkIIF(SortMethod="itemidAS","��",""))%></td>
			<td rowspan="2">�̹���</td>
			<td width="70" align="center" colspan="2">��</td>
			<td width="70" align="center" colspan="2">PC��</td>
			<td width="70" align="center" colspan="2">�������</td>
			<td width="70" align="center" colspan="2">APP</td>
			<td width="70" align="center" colspan="2">���޸�</td>
			<td width="70" align="center" rowspan="2" onClick="chgSortMethod('<%=chkIIF(SortMethod="wishDS","wishAS","wishDS")%>')" style="cursor:pointer;">Wish<%=chkIIF(SortMethod="wishDS","��",chkIIF(SortMethod="wishAS","��",""))%></td>
			<td width="70" align="center" rowspan="2">�󼼺���</td>
		</tr>
		<tr bgcolor="#EDEDFF">
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="totPrcDS","totPrcAS","totPrcDS")%>')" style="cursor:pointer;">�Ǹž�<%=chkIIF(SortMethod="totPrcDS","��",chkIIF(SortMethod="totPrcAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="totNoDS","totNoAS","totNoDS")%>')" style="cursor:pointer;">�ǸŰ���<%=chkIIF(SortMethod="totNoDS","��",chkIIF(SortMethod="totNoAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="pcPrcDS","pcPrcAS","pcPrcDS")%>')" style="cursor:pointer;">�Ǹž�<%=chkIIF(SortMethod="pcPrcDS","��",chkIIF(SortMethod="pcPrcAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="pcNoDS","pcNoAS","pcNoDS")%>')" style="cursor:pointer;">�ǸŰ���<%=chkIIF(SortMethod="pcNoDS","��",chkIIF(SortMethod="pcNoAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="mobPrcDS","mobPrcAS","mobPrcDS")%>')" style="cursor:pointer;">�Ǹž�<%=chkIIF(SortMethod="mobPrcDS","��",chkIIF(SortMethod="mobPrcAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="mobNoDS","mobNoAS","mobNoDS")%>')" style="cursor:pointer;">�ǸŰ���<%=chkIIF(SortMethod="mobNoDS","��",chkIIF(SortMethod="mobNoAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="appPrcDS","appPrcAS","appPrcDS")%>')" style="cursor:pointer;">�Ǹž�<%=chkIIF(SortMethod="appPrcDS","��",chkIIF(SortMethod="appPrcAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="appNoDS","appNoAS","appNoDS")%>')" style="cursor:pointer;">�ǸŰ���<%=chkIIF(SortMethod="appNoDS","��",chkIIF(SortMethod="appNoAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="extPrcDS","extPrcAS","extPrcDS")%>')" style="cursor:pointer;">�Ǹž�<%=chkIIF(SortMethod="extPrcDS","��",chkIIF(SortMethod="extPrcAS","��",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="extNoDS","extNoAS","extNoDS")%>')" style="cursor:pointer;">�ǸŰ���<%=chkIIF(SortMethod="extNoDS","��",chkIIF(SortMethod="extNoAS","��",""))%></td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
			<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oReport.FMasterItemList(i).FItemid %>" target="_blank" title="�̸�����"><%= oReport.FMasterItemList(i).FItemid %></a></td>
			<td><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(oReport.FMasterItemList(i).FItemid)%>/<%=oReport.FMasterItemList(i).Fsmallimage%>"></td>
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
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=D&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">����</a></td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "O"  '// �ɼǺ� �̺�Ʈ ���
		call oReport.GetEventStatisticsByItemOptionDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td width="90" align="center">�����۹�ȣ</td>
			<td width="90" align="center">�ɼǹ�ȣ</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="500" align="center">�׷���</td>
			<td width="70" align="center">�󼼺���</td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><table class="a"><tr><td><%= oReport.FMasterItemList(i).FItemid %></td><td><div id="imgview<%=i%>"><span onClick="viewImage('imgview<%=i%>','<%= oReport.FMasterItemList(i).FItemid %>')" style="cursor:pointer">[view]</span></div></td></tr></table></td>
			<td align="center"><%= oReport.FMasterItemList(i).FItemOption %></td>
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
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=D&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&ItemOption=<%= oReport.FMasterItemList(i).FItemOption %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">����</a></td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "M"  '// �귣�庰 �̺�Ʈ ���
		call oReport.GetEventStatisticsByMakerIDDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td width="150" align="center">�귣��</td>
			<td width="70" align="center">�Ǹž�</td>
			<td width="70" align="center">�ǸŰ���</td>
			<td width="500" align="center">�׷���</td>
			<td width="70" align="center">�󼼺���</td>
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
			<td align="center">
				<a href="/admin/report/event_report_detail.asp?SType=T&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&makerid=<%= oReport.FMasterItemList(i).Fmakerid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">����</a>
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
			<td align="right"><%= FormatNumber(t_TotalCost,0) %></td>
			<td align="right"><%= FormatNumber(t_FTotalNo,0) %> ��</td>
		</tr>
	</table>

<%
set oReport = Nothing
'set oTotal = Nothing
%>
<iframe src="about:blank" name="iframeDB1" width="0" height="0">
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
