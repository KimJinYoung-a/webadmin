<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이벤트통계
' Hieditor : 서동석 생성
'			 2021.02.23 한용민 수정(검색조건 추가. 최근디자인으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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
	sKind = requestCheckVar(Request("eventkind"),10)	'이벤트종류
	eType = requestCheckVar(Request("eventtype"),10)	'이벤트종류
	dispCate	= requestCheckVar(Request("disp"),10) 		'전시 카테고리
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

Call fnSetEventCommonCode '공통코드 어플리케이션 변수에 세팅

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
		response.write "한번에 2달 이상을 검색할 수 없습니다."
		dbget.close()
		response.end
	end if

	oReport.GetEventStatisticsDataMart

'oReport.GetEventStatisticsAll
'IF ReportType="s" Then
''oReport.GetEventStatisticsAllSelectedTerm
'ELSE
''oReport.OLD_GetEventStatisticsAll
'End IF
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function changecontent(){
	document.frm.target="";
	document.frm.action="";
	document.frm.submit();
}

//리스트 정렬
function jstrSort(sValue,i){
 	document.frm.selSort.value= sValue;

	   if (-1 < eval("document.frmList.img"+i).src.indexOf("_alpha")){
        document.frm.selSort.value= sValue+"D";
    }else if (-1 < eval("document.frmList.img"+i).src.indexOf("_bot")){
     		document.frm.selSort.value= sValue+"A";
    }else{
       document.frm.selSort.value= sValue+"D";
    }
	 document.frm.target="";
	 document.frm.action="";
	 document.frm.submit();
}

// 엑셀받기
function fnGetExcelFile() {
	document.frm.target="_blank";
	document.frm.action="/admin/report/event_report2_excel.asp";
	document.frm.submit();
}

</script>

<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="selSort" value="<%=strSort%>"><!--정렬-->
<input type="hidden" name="reloading" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		<table class="a" border="0" cellpadding="3">
		<tr>
		<td class="a" >
			<!--
			<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
			-->

			* 기간:
				<input type="text" name="Sdate" value="<%=Sdate%>" size="10" readonly onclick="jsPopCal('Sdate');">~
				<input type="text" name="Edate" value="<%=Edate%>" size="10" readonly onclick="jsPopCal('Edate');">
				<input type="radio" name="rt" value="s" <% IF ReportType="s" Then response.write "checked" %>>선택 기간별 매출
			<input type="radio" name="rt" value="e" <% IF ReportType="e" Then response.write "checked" %>>이벤트 기간별 매출

			: (약 1시간 지연 데이터)
		</td>
	</tr>
	<tr>
		<td>
			* 이벤트 종류 <%sbGetOptEventCodeValue "eventkind", sKind, True,""%>
			&nbsp;
			* 이벤트 유형 <%sbGetOptCommonCodeArr "eventtype", eType, True,True,"onChange='changecontent();'"%>
			&nbsp;
			* 이벤트 번호 : <input type="text" size="10" name="eventid" value="<%=eventid%>">
			&nbsp;
			* 관리카테고리: <% DrawSelectBoxCategoryLarge "cateNo",cateNo %>
			&nbsp;
			* 전시카테고리: <%=fnDispCateSelectBox(1,"","disp",dispCate,"") %>
		</td>
	</tr>
	</table>
</td>
	<td class="a" align="center" bgcolor="<%= adminColor("gray") %>"><a href="javascript:changecontent();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
</tr>
</table>
</form>
<br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
		<img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" onclick="fnGetExcelFile()" style="cursor:pointer" />
		<!--정렬
				<select name="selSort" class="select" onChange="javascript:document.frm.submit();">
					<option value="1" <%if strSort="1" then%>selected<%end if%>>이벤트코드순</option>
					<option value="2" <%if strSort="2" then%>selected<%end if%>>매출순</option>
					<option value="3" <%if strSort="3" then%>selected<%end if%>>수익순</option>
				</select>-->
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= oReport.FResultCount %></b>
		<% if oReport.FResultCount > 0 then %>
			&nbsp;
			총이벤트매출액 :
			<%
				ttSellPrice = 0
				for i=0 to oReport.FResultCount-1
					ttSellPrice = ttSellPrice + oReport.FMasterItemList(i).Fselltotal
				next
				Response.Write FormatNumber(ttSellPrice,0)
			%>원 /
			총평균매출액 : <%=FormatNumber(ttSellPrice/oReport.FResultCount,0) %>원
		<% end if %>
	</td>
</tr>

<form name="frmList" method="post" style="margin:0px;" >
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60" rowspan="2" onClick="javascript:jstrSort('E','1');" style="cursor:hand;"><b>이벤트<br>번호</b><img src="/images/list_lineup<%IF strSort="ED" THEN%>_bot<%ELSEIF strSort="EA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
	<td rowspan="2">이벤트명</td>
	<td rowspan="2">이벤트 기간</td>
	<td  colspan="5">Mobile/App</td>
	<td colspan="5"> PC-Web </td>
	<td colspan="5">제휴</td>
	<td colspan="5">3PL</td>
	<td  rowspan="2" >총 판매수</td>
	<td  rowspan="2" onClick="javascript:jstrSort('TM','2');" style="cursor:hand;"><b>매출합계</b><img src="/images/list_lineup<%IF strSort="TMD" THEN%>_bot<%ELSEIF strSort="TMA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
	<td  rowspan="2" onClick="javascript:jstrSort('TR','12');" style="cursor:hand;"><b>취급액</b><img src="/images/list_lineup<%IF strSort="TRD" THEN%>_bot<%ELSEIF strSort="TRA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img12"></td>
	<td  rowspan="2" onClick="javascript:jstrSort('TP','3');" style="cursor:hand;"><b>수익</b><img src="/images/list_lineup<%IF strSort="TPD" THEN%>_bot<%ELSEIF strSort="TPA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
	<td width="150" rowspan="2">상세</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>판매수</td>
	<td onClick="javascript:jstrSort('MM','4');" style="cursor:hand;"><b>매출</b> <img src="/images/list_lineup<%IF strSort="MMD" THEN%>_bot<%ELSEIF strSort="MMA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
	<td onClick="javascript:jstrSort('MR','13');" style="cursor:hand;"><b>취급액</b> <img src="/images/list_lineup<%IF strSort="MRD" THEN%>_bot<%ELSEIF strSort="MRA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img13"></td>
	<td width=40>점유율</td>
	<td onClick="javascript:jstrSort('MP','5');" style="cursor:hand;"><b>수익</b> <img src="/images/list_lineup<%IF strSort="MPD" THEN%>_bot<%ELSEIF strSort="MPA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
	<td>판매수</td>
	<td onClick="javascript:jstrSort('WM','6');" style="cursor:hand;"><b>매출</b> <img src="/images/list_lineup<%IF strSort="WMD" THEN%>_bot<%ELSEIF strSort="WMA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img6"></td>
	<td onClick="javascript:jstrSort('WR','14');" style="cursor:hand;"><b>취급액</b> <img src="/images/list_lineup<%IF strSort="WRD" THEN%>_bot<%ELSEIF strSort="WRA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img14"></td>
	<td width=40>점유율</td>
	<td onClick="javascript:jstrSort('WP','7');" style="cursor:hand;"><b>수익</b> <img src="/images/list_lineup<%IF strSort="WPD" THEN%>_bot<%ELSEIF strSort="WPA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img7"></td>
	<td>판매수</td>
	<td onClick="javascript:jstrSort('BM','8');" style="cursor:hand;"><b>매출</b> <img src="/images/list_lineup<%IF strSort="BMD" THEN%>_bot<%ELSEIF strSort="BMA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img8"></td>
	<td onClick="javascript:jstrSort('BR','15');" style="cursor:hand;"><b>취급액</b> <img src="/images/list_lineup<%IF strSort="BRD" THEN%>_bot<%ELSEIF strSort="BRA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img15"></td>
	<td width=40>점유율</td>
	<td onClick="javascript:jstrSort('BP','9');" style="cursor:hand;"><b>수익</b> <img src="/images/list_lineup<%IF strSort="BPD" THEN%>_bot<%ELSEIF strSort="BPA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img9"></td>
	<td>판매수</td>
	<td onClick="javascript:jstrSort('3M','10');" style="cursor:hand;"><b>매출</b> <img src="/images/list_lineup<%IF strSort="3MD" THEN%>_bot<%ELSEIF strSort="3MA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img10"></td>
	<td onClick="javascript:jstrSort('3R','16');" style="cursor:hand;"><b>취급액</b> <img src="/images/list_lineup<%IF strSort="3RD" THEN%>_bot<%ELSEIF strSort="3RA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img16"></td>
	<td width=40>점유율</td>
	<td onClick="javascript:jstrSort('3P','11');" style="cursor:hand;"><b>수익</b> <img src="/images/list_lineup<%IF strSort="3PD" THEN%>_bot<%ELSEIF strSort="3PA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img11"></td>
</tr>
<tr bgcolor="#EEEEEE"  align="center">
	<td colspan="3" align="center">총합계</td>
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
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFFF';>
	<td align="center"><a href="<%= wwwURL %>/event/eventmain.asp?eventid=<%= oReport.FMasterItemList(i).FEventIdx %>" target="_blank"><%= oReport.FMasterItemList(i).FEventIdx %></a></td>
	<!--<td align="center">
	<% if Not(oReport.FMasterItemList(i).FEventBanImage="" or isNull(oReport.FMasterItemList(i).FEventBanImage)) then %>
		<img src="<%= oReport.FMasterItemList(i).FEventBanImage %>" height="42" align="absmiddle">
	<% else %>
		<img src="http://fiximage.10x10.co.kr/images/spacer.gif" height="42" align="absmiddle">
	<% end if %>
	</td-->
	<td align="left">
		<a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oReport.FMasterItemList(i).FEventIdx %>" target="_blank">
		<%= oReport.FMasterItemList(i).FEventName %></a>
	</td>
	<td align="center">
		<%= oReport.FMasterItemList(i).FStartDay %>~<%= oReport.FMasterItemList(i).FEndDay %>
	</td>
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
	<td align="center" style="word-wrap:break-word;word-break:break-all;white-space:nowrap;">
		<a href="/admin/report/event_report_detail.asp?SType=D&eventid=<%= oReport.FMasterItemList(i).FEventIdx %>&SDate=<%=oReport.FMasterItemList(i).FStartDay%>&EDate=<%= oReport.FMasterItemList(i).FEndDay %>" target="_blank">날짜별</a>
		|
		<a href="/admin/report/event_report_detail.asp?SType=T&eventid=<%= oReport.FMasterItemList(i).FEventIdx %>&SDate=<%=oReport.FMasterItemList(i).FStartDay%>&EDate=<%= oReport.FMasterItemList(i).FEndDay %>" target="_blank">상품별</a>
		|
		<a href="/admin/report/event_report_detail.asp?SType=M&eventid=<%= oReport.FMasterItemList(i).FEventIdx %>&SDate=<%=oReport.FMasterItemList(i).FStartDay%>&EDate=<%= oReport.FMasterItemList(i).FEndDay %>" target="_blank">브랜드별</a>
	</td>
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="30" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</form>

<%
set oReport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
