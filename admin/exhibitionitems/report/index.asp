<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionReportCls.asp"-->
<%
dim Sdate, Edate, page , i
dim mastercode , detailcode
dim viewtype

Sdate = requestCheckVar(request("Sdate"),10)
Edate = requestCheckVar(request("Edate"),10)
mastercode = requestCheckVar(request("mastercode"),10)
detailcode = requestCheckVar(request("detailcode"),10)
viewtype = requestCheckVar(request("viewtype"),1)

IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

if mastercode = "" then mastercode = 0
if detailcode = "" then detailcode = 0
if viewtype = "" then viewtype = "E"


dim oReport
set oReport = new ExhibitionReport
oReport.FRectStart      = Sdate
oReport.FRectEnd        = dateSerial(year(Edate),month(EDate),Day(EDate)+1)
oReport.FrectMasterCode = mastercode
oReport.FrectDetailCode = detailcode
if viewtype = "E" then '// 기획전
    oReport.GetExhibitionStatisticsDataMart
else '// 하위이벤트
    oReport.GetSubEventStatisticsTotalDataMart
end if 

%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css"/>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script>
    function changecontent(){
		document.frm.target="";
		document.frm.action="";
		document.frm.submit();
    }

    // 상세보기 팝업
    function popReportDetail(type,mastercode,detailcode,sdate,edate){
        var popReportDetail = window.open('/admin/exhibitionitems/report/pop_event_report_detail.asp?SType='+type+'&mastercode='+mastercode+'&detailcode='+detailcode+'&SDate='+sdate+'&EDate='+edate,'popReportDetail','width=1024,height=768,resizable=yes,scrollbars=yes')
        popReportDetail.focus();
    }
</script>
<div class="content scrl" style="top:40px;">
	<div class="pad20">
		<!-- 상단 검색폼 시작 -->
		<div>
            <form name="frm" method="get" action="">
            <input type="hidden" name="page" value="1">
            <input type="hidden" name="menupos" value="<%= request("menupos") %>">
                <table class="tbType1 listTb">
                    <tr>
                        <td width="70" bgcolor="<%= adminColor("gray") %>">검색조건</td>
                        <td style="text-align:left">
                            지정일자 <input id="SDate" name="SDate" value="<%=Sdate%>" class="text" size="10" maxlength="10"/><img src="http://scm.10x10.co.kr/images/calicon.gif" id="SDate_trigger" border="0" style="cursor:pointer;vertical-align:middle;"/>
                            ~ <input id="Edate" name="Edate" value="<%=Edate%>" class="text" size="10" maxlength="10"/><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="Edate_trigger" border="0" style="cursor:pointer;vertical-align:middle;"/>
                            <script type="text/javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "SDate", trigger    : "SDate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
                            <script type="text/javascript">
                                var CAL_End = new Calendar({
                                    inputField : "Edate", trigger    : "Edate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
                            <br/>
                            <br/>
                            목록 <% DrawMainPosCodeCombo "mastercode", mastercode ,"" %>
                            <% if mastercode > 0 then %>
                                <% DrawDetailSelectBox "detailcode" , detailcode , mastercode %>
                            <% end if %>
                            <br/>
                            <br/>
                            <input type="radio" name="viewtype" value="E" <%=chkiif(viewtype="E","checked","")%> id="exhibition"/> : <label for="exhibition">기획전 </label>
                            <input type="radio" name="viewtype" value="S" <%=chkiif(viewtype="S","checked","")%> id="subevent"/> : <label for="subevent">하위이벤트 </label>
                        </td>
                        <td width="50" bgcolor="<%= adminColor("gray") %>">
                            <input type="button" class="button_s" value="검색" onClick="javascript:changecontent();">
                        </td>
                    </tr>
                </table>
            </form>
        </div>

        <div class="tPad15">
            <% if oReport.FResultCount > 0 then %>
            <table width="100%" cellspacing="0" cellpadding="3" class="a">
                <tr>
                    <td>총기획전 수 : <%=oReport.FResultCount%>개</td>
                    <td align="right">
                        총기획전 매출액 :
                        <%
                            dim ttSellPrice : ttSellPrice = 0
                            for i=0 to oReport.FResultCount-1
                                ttSellPrice = ttSellPrice + oReport.ExhibitionReportList(i).Fselltotal
                            next
                            Response.Write FormatNumber(ttSellPrice,0)
                        %>원 /
                        총평균 매출액 : <%=FormatNumber(ttSellPrice/oReport.FResultCount,0) %>원
                    </td>
                </tr>
            </table>
        </div>
        <div class="tPad15">
            <form name="frmList" method="post">
            <table class="tbType1 listTb">
                <tr bgcolor="#DDDDFF" align="center">
                    <td width="150" rowspan="2"><b>기획전명</b></td>
                    <td width="100" rowspan="2">하위<br/>카테고리</td>
                    <% IF viewtype="S" THEN %>
                    <td width="40" rowspan="2"><b>이벤트번호</b></td>
		            <td rowspan="2">이벤트명<br/>시작일/종료일</td>
                    <% END IF %>
                    <td colspan="4">Mobile/App</td>
                    <td colspan="4"> PC-Web </td>
                    <td colspan="4">제휴</td>
                    <td colspan="4">3PL</td>
                    <td rowspan="2">총 판매수</td>
                    <td rowspan="2"><b>매출합계</b></td>
                    <td rowspan="2"><b>수익</b></td>
                    <td width="150" rowspan="2">상세 보기 </td>
                </tr>
                <tr bgcolor="#DDDDFF" align="center">
                    <td>판매수</td>
                    <td>매출</td>
                    <td>점유율</td>
                    <td>수익</td>
                    <td>판매수</td>
                    <td>매출</td>
                    <td>점유율</td>
                    <td>수익</td>
                    <td>판매수</td>
                    <td>매출</td>
                    <td>점유율</td>
                    <td>수익</td>
                    <td>판매수</td>
                    <td>매출</td>
                    <td>점유율</td>
                    <td>수익</td>
                </tr>
                <tr bgcolor="#EEEEEE"  align="right">
                    <td colspan="<%=chkiif(viewtype="E","2","4")%>" align="center">총합계</td>
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

                    <td colspan="2"></td>
                </tr>
                <%
                    for i=0 to oReport.FResultCount-1
                %>
                <tr bgcolor="#FFFFFF" align="right">
                    <td align="center">
                        <%=getMasterCodeName(oReport.ExhibitionReportList(i).Fmastercode)%>
                    </td>
                    <td align="left">
                        <% if oReport.ExhibitionReportList(i).Fdetailcode > 0 then %>
						    <%=getDetailCodeName(oReport.ExhibitionReportList(i).Fmastercode,oReport.ExhibitionReportList(i).Fdetailcode)%>
                        <% else %>
                            <span style="color:#ed121d">기획전 최상위</span>
						<% end if %>
                    </td>
                    
                    <% IF viewtype = "S" THEN %>
                    <td><%=oReport.ExhibitionReportList(i).Fevt_code%></td>
                    <td><%=oReport.ExhibitionReportList(i).Fevt_name%><br/><strong><span style="color:red"><%=oReport.ExhibitionReportList(i).Fevt_startdate%> - <%=oReport.ExhibitionReportList(i).Fevt_enddate%></span></strong></td>
                    <% end if %>

                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_mobile,0) %></td>
                    <td bgcolor="#DDFFDD"><b><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_mobile,0) %></b></td>
                    <td bgcolor="#DDFFDD"><b><%if oReport.ExhibitionReportList(i).Fsellsum_mobile > 0 and oReport.ExhibitionReportList(i).Fselltotal > 0 then %><%= FormatNumber((oReport.ExhibitionReportList(i).Fsellsum_mobile/oReport.ExhibitionReportList(i).Fselltotal)*100,0) %>%<%end if%></b></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_mobile -oReport.ExhibitionReportList(i).Fbuysum_mobile,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_PC,0) %></td>
                    <td bgcolor="#EEEEEE"><b><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_PC,0) %></b></td>
                    <td bgcolor="#EEEEEE"><b><%if  oReport.ExhibitionReportList(i).Fsellsum_PC > 0 and oReport.ExhibitionReportList(i).Fselltotal > 0 then %> <%=FormatNumber((oReport.ExhibitionReportList(i).Fsellsum_PC/oReport.ExhibitionReportList(i).Fselltotal)*100,0)%>%<%end if%></b></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_PC -oReport.ExhibitionReportList(i).Fbuysum_PC,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_outmall,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_outmall,0) %></td>
                    <td><%if oReport.ExhibitionReportList(i).Fsellsum_outmall > 0 and oReport.ExhibitionReportList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_outmall/oReport.ExhibitionReportList(i).Fselltotal*100,0) %>%<%end if%></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_outmall -oReport.ExhibitionReportList(i).Fbuysum_outmall,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_3PL,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_3PL,0) %></td>
                    <td><%if oReport.ExhibitionReportList(i).Fsellsum_3PL > 0 and oReport.ExhibitionReportList(i).Fselltotal > 0 then %><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_3PL/oReport.ExhibitionReportList(i).Fselltotal*100,0) %>%<%end if%></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_3PL -oReport.ExhibitionReportList(i).Fbuysum_3PL,0) %></td>
                    <td><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt,0) %></td>
                    <td><b><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></b></td>
                    <td><b><%=FormatNumber(oReport.ExhibitionReportList(i).Fselltotal-oReport.ExhibitionReportList(i).Fbuytotal,0)%></b></td>
                    <% if viewtype="E" then %>
                    <td align="center">
                        <a href="" onclick="popReportDetail('D','<%=oReport.ExhibitionReportList(i).Fmastercode%>','<%=oReport.ExhibitionReportList(i).Fdetailcode%>','<%=Sdate%>','<%=Edate%>');return false;" target="_blank">날짜</a>
                        |
                        <a href="" onclick="popReportDetail('T','<%=oReport.ExhibitionReportList(i).Fmastercode%>','<%=oReport.ExhibitionReportList(i).Fdetailcode%>','<%=Sdate%>','<%=Edate%>');return false;" target="_blank">상품</a>
                        |
                        <a href="" onclick="popReportDetail('M','<%=oReport.ExhibitionReportList(i).Fmastercode%>','<%=oReport.ExhibitionReportList(i).Fdetailcode%>','<%=Sdate%>','<%=Edate%>');return false;" target="_blank">브랜드</a>
                    </td>
                    <% else %>
                    <td align="center">
                        <a href="/admin/report/event_report_detail.asp?SType=D&eventid=<%=oReport.ExhibitionReportList(i).Fevt_code%>&SDate=<%=oReport.ExhibitionReportList(i).Fevt_startdate%>&EDate=<%=oReport.ExhibitionReportList(i).Fevt_enddate%>" target="_blank">날짜</a>
                        |
                        <a href="/admin/report/event_report_detail.asp?SType=T&eventid=<%=oReport.ExhibitionReportList(i).Fevt_code%>&SDate=<%=oReport.ExhibitionReportList(i).Fevt_startdate%>&EDate=<%=oReport.ExhibitionReportList(i).Fevt_enddate%>" target="_blank">상품</a>
                        |
                        <a href="/admin/report/event_report_detail.asp?SType=M&eventid=<%=oReport.ExhibitionReportList(i).Fevt_code%>&SDate=<%=oReport.ExhibitionReportList(i).Fevt_startdate%>&EDate=<%=oReport.ExhibitionReportList(i).Fevt_enddate%>" target="_blank">브랜드</a>
                    </td>
                    <% end if %>
                </tr>
                <%
                    next
                %>
            </table>
            </form>
            <% else %>
            <table class="tbType1 listTb">
                <tr bgcolor="#DDDDFF">
                    <td align="center">[결과가 없습니다]
                    </td>
                </tr>
            </table>
            <% end if %>
            </div>
        </div>
	</div>
</div>
<%
    set oReport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
