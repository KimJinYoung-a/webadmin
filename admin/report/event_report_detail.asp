<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<%

dim SType '// 분류
dim EventID,ItemID, itemoption,i, makerid
dim BasicDateSet, Sdate, Edate, page, grpWidth
dim sortMethod

Dim oldlist


SType = requestCheckVar(request("SType"),10)
EventID = requestCheckVar(request("EventID"),10)
ItemID = requestCheckVar(request("ItemID"),10)
itemoption = requestCheckVar(request("itemoption"),10)  ''2013/10/14 추가
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




dim  oReport  '// 통계 데이타
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

'dim oTotal '// 총합계 ?? 필요?
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

	// 엑셀받기
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
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		-->
		검색 기간 :
			<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
			<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
		<br />

		이벤트 번호 :
			<input type="text" name="EventID" size="10" value="<%= EventID %>">
        브랜드 :
			<input type="text" name="makerid" size="10" value="<%= makerid %>">
        상품 번호 :
            <input type="text" name="ITEMID" size="9" value="<%= ITEMID %>">
        옵션 번호 :
            <input type="text" name="itemoption" size="9" value="<%= itemoption %>">
		<br />
		분류 :
			<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>> 날짜별
			<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>> 상품별
			<input type="radio" name="SType" value="O" <% If SType = "O" Then response.write "checked" %>> 옵션별
			<input type="radio" name="SType" value="M" <% If SType = "M" Then response.write "checked" %>> 브랜드별
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

	CASE "D" '// 날짜별 이벤트 통계
	    IF (ItemID<>"") then
	        call oReport.GetEventStatisticsByDate
		ELSE
		    call oReport.GetEventStatisticsByDateDataMart
		END IF
%>
		<tr bgcolor="#DDDDFF">
	    	<td width="90" align="center">구매일</td>
	    	<td width="70" align="center">판매액</td>
			<td width="70" align="center">판매갯수</td>
			<td width="500" align="center">그래프</td>
			<td width="70" align="center">상세보기</td>
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
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>개</td>
			<td width="500">
				<%
					'그래프 길이 계산 (2008.07.08;허진원 수정)
					if oReport.maxc>0 then
						grpWidth = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxc*400)
					else
					grpWidth = 0
					end if
				%>
				<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
			</td>
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=T&EventID=<%= EventID %>&SDate=<%=oReport.FMasterItemList(i).Fselldate%>&EDate=<%= oReport.FMasterItemList(i).Fselldate %>">보기</a></td>
   </tr>
		<% next %>
	<% end if %>

<%
	CASE "T"  '// 상품별 이벤트 통계
		oReport.FRectSort = sortMethod
		call oReport.GetEventStatisticsByItemIDDataMart
%>
		<tr bgcolor="#EDEDFF">
			<td width="150" align="center" rowspan="2">브랜드</td>
			<td width="90" align="center" rowspan="2" onClick="chgSortMethod('<%=chkIIF(SortMethod="itemidDS","itemidAS","itemidDS")%>')" style="cursor:pointer;">아이템번호<%=chkIIF(SortMethod="itemidDS","▼",chkIIF(SortMethod="itemidAS","▲",""))%></td>
			<td rowspan="2">이미지</td>
			<td width="70" align="center" colspan="2">계</td>
			<td width="70" align="center" colspan="2">PC웹</td>
			<td width="70" align="center" colspan="2">모바일웹</td>
			<td width="70" align="center" colspan="2">APP</td>
			<td width="70" align="center" colspan="2">제휴몰</td>
			<td width="70" align="center" rowspan="2" onClick="chgSortMethod('<%=chkIIF(SortMethod="wishDS","wishAS","wishDS")%>')" style="cursor:pointer;">Wish<%=chkIIF(SortMethod="wishDS","▼",chkIIF(SortMethod="wishAS","▲",""))%></td>
			<td width="70" align="center" rowspan="2">상세보기</td>
		</tr>
		<tr bgcolor="#EDEDFF">
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="totPrcDS","totPrcAS","totPrcDS")%>')" style="cursor:pointer;">판매액<%=chkIIF(SortMethod="totPrcDS","▼",chkIIF(SortMethod="totPrcAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="totNoDS","totNoAS","totNoDS")%>')" style="cursor:pointer;">판매갯수<%=chkIIF(SortMethod="totNoDS","▼",chkIIF(SortMethod="totNoAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="pcPrcDS","pcPrcAS","pcPrcDS")%>')" style="cursor:pointer;">판매액<%=chkIIF(SortMethod="pcPrcDS","▼",chkIIF(SortMethod="pcPrcAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="pcNoDS","pcNoAS","pcNoDS")%>')" style="cursor:pointer;">판매갯수<%=chkIIF(SortMethod="pcNoDS","▼",chkIIF(SortMethod="pcNoAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="mobPrcDS","mobPrcAS","mobPrcDS")%>')" style="cursor:pointer;">판매액<%=chkIIF(SortMethod="mobPrcDS","▼",chkIIF(SortMethod="mobPrcAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="mobNoDS","mobNoAS","mobNoDS")%>')" style="cursor:pointer;">판매갯수<%=chkIIF(SortMethod="mobNoDS","▼",chkIIF(SortMethod="mobNoAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="appPrcDS","appPrcAS","appPrcDS")%>')" style="cursor:pointer;">판매액<%=chkIIF(SortMethod="appPrcDS","▼",chkIIF(SortMethod="appPrcAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="appNoDS","appNoAS","appNoDS")%>')" style="cursor:pointer;">판매갯수<%=chkIIF(SortMethod="appNoDS","▼",chkIIF(SortMethod="appNoAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="extPrcDS","extPrcAS","extPrcDS")%>')" style="cursor:pointer;">판매액<%=chkIIF(SortMethod="extPrcDS","▼",chkIIF(SortMethod="extPrcAS","▲",""))%></td>
			<td width="70" align="center" onClick="chgSortMethod('<%=chkIIF(SortMethod="extNoDS","extNoAS","extNoDS")%>')" style="cursor:pointer;">판매갯수<%=chkIIF(SortMethod="extNoDS","▼",chkIIF(SortMethod="extNoAS","▲",""))%></td>
		</tr>
	<% if oReport.FResultCount > 0 then %>
		<% for i=0 to oReport.FResultCount-1 %>
		<%
		t_TotalCost = t_TotalCost + oReport.FMasterItemList(i).Fselltotal
		t_FTotalNo  = t_FTotalNo + oReport.FMasterItemList(i).Fsellcnt
		%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%= oReport.FMasterItemList(i).Fmakerid %></td>
			<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oReport.FMasterItemList(i).FItemid %>" target="_blank" title="미리보기"><%= oReport.FMasterItemList(i).FItemid %></a></td>
			<td><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(oReport.FMasterItemList(i).FItemid)%>/<%=oReport.FMasterItemList(i).Fsmallimage%>"></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fselltotal,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt,0) %>개</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_PC,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_PC,0) %>개</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_mobile,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_mobile,0) %>개</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_App,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_App,0) %>개</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellsum_outmall,0) %></td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).Fsellcnt_outmall,0) %>개</td>
			<td align="right"><%= FormatNumber(oReport.FMasterItemList(i).FwishCnt,0) %>건</td>
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=D&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">보기</a></td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "O"  '// 옵션별 이벤트 통계
		call oReport.GetEventStatisticsByItemOptionDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td width="90" align="center">아이템번호</td>
			<td width="90" align="center">옵션번호</td>
			<td width="70" align="center">판매액</td>
			<td width="70" align="center">판매갯수</td>
			<td width="500" align="center">그래프</td>
			<td width="70" align="center">상세보기</td>
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
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>개</td>
			<td>
				<%
				'그래프 길이 계산 (2008.07.08;허진원 수정)
					if oReport.maxc>0 then
						grpWidth = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxc*400)
					else
						grpWidth = 0
					end if
				%>
				<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
			</td>
			<td align="center"><a href="/admin/report/event_report_detail.asp?SType=D&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&ItemOption=<%= oReport.FMasterItemList(i).FItemOption %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">보기</a></td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE "M"  '// 브랜드별 이벤트 통계
		call oReport.GetEventStatisticsByMakerIDDataMart
%>
		<tr bgcolor="#DDDDFF">
			<td width="150" align="center">브랜드</td>
			<td width="70" align="center">판매액</td>
			<td width="70" align="center">판매갯수</td>
			<td width="500" align="center">그래프</td>
			<td width="70" align="center">상세보기</td>
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
			<td align="right"><%= oReport.FMasterItemList(i).Fsellcnt %>개</td>
			<td>
				<%
					'그래프 길이 계산 (2008.07.08;허진원 수정)
					if oReport.maxc>0 then
						grpWidth = Clng(oReport.FMasterItemList(i).Fselltotal/oReport.maxc*400)
					else
						grpWidth = 0
					end if
				%>
				<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%=grpWidth%>">
			</td>
			<td align="center">
				<a href="/admin/report/event_report_detail.asp?SType=T&EventID=<%= EventID %>&ItemID=<%= oReport.FMasterItemList(i).FItemid %>&makerid=<%= oReport.FMasterItemList(i).Fmakerid %>&SDate=<%=Sdate%>&EDate=<%=Edate%>">보기</a>
			</td>
		</tr>
		<% next %>
	<% end if %>
<%
	CASE ELSE
		response.write "오류발생,다시 시도"
END SELECT
%>
		<tr>
			<td align="center">총합</td>
			<td align="right"><%= FormatNumber(t_TotalCost,0) %></td>
			<td align="right"><%= FormatNumber(t_FTotalNo,0) %> 개</td>
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
