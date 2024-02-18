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
dim SType '// 분류
dim makerid , mastercode , detailcode , Sdate , Edate
dim i , grpWidth

SType = requestCheckVar(request("SType"),10)
makerid = requestCheckVar(request("makerid"),32)
mastercode = requestCheckVar(request("mastercode"),10)
detailcode = requestCheckVar(request("detailcode"),10)
Sdate = requestCheckVar(request("SDate"),10)
Edate = requestCheckVar(request("EDate"),10)

IF Sdate="" THEN
	Sdate= dateSerial(Year(now()),Month(now()),day(now()))
End IF

IF Edate="" THEN
	Edate= dateSerial(Year(now()),Month(now()),day(now())+1)
End IF

if mastercode = "" then mastercode = 0
if detailcode = "" then detailcode = 0

dim oReport  '// 통계 데이타
set oReport = new ExhibitionReport
	oReport.FRectMakerid = makerid
	oReport.FRectStart = Sdate
	oReport.FRectEnd = dateSerial(year(Edate),month(EDate),Day(EDate))
	oReport.FrectMasterCode = mastercode
	oReport.FrectDetailCode = detailcode

dim t_TotalCost, t_FTotalNo
t_TotalCost = 0
t_FTotalNo  = 0
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
<script type="text/javascript">
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}

	function viewImage(div,itemid)
	{
		iframeDB1.location.href = "/admin/report/iframe_viewImage.asp?div="+div+"&itemid="+itemid+"";
	}
</script>

<div class="content scrl" style="top:40px;">
	<div class="pad20">
		<div>
			<h1>기획전통계 날짜 , 상품 , 브랜드 상세보기</h1>
		<div>
		<!-- 상단 검색폼 시작 -->
		<div class="tPad15">
			<form name="frm" method="get" action="">
			<table class="tbType1 listTb">
			<input type="hidden" name="page" value="1">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
				<tr>
					<td width="70" bgcolor="<%= adminColor("gray") %>">검색조건</td>
					<td style="text-align:left">
					검색기간 : <input id="SDate" name="SDate" value="<%=Sdate%>" class="text" size="10" maxlength="10"/><img src="http://scm.10x10.co.kr/images/calicon.gif" id="SDate_trigger" border="0" style="cursor:pointer;vertical-align:middle;"/>
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
					<br/><br/>
					목록 : <% DrawMainPosCodeCombo "mastercode", mastercode ,"" %>
					<% if mastercode > 0 then %>
						<% DrawDetailSelectBox "detailcode" , detailcode , mastercode %>
					<% end if %>
					<br/><br/>
					분류 :
						<input type="radio" name="SType" value="D" <% If SType = "D" Then response.write "checked" %>> 날짜별
						<input type="radio" name="SType" value="T" <% If SType = "T" Then response.write "checked" %>> 상품별
						<input type="radio" name="SType" value="M" <% If SType = "M" Then response.write "checked" %>> 브랜드별
					</td>
					<td width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
					</td>
				</tr>
			</table>
			</form>
		</div>

		<div class="tPad15">
			<table class="tbType1 listTb">
				<%
				SELECT CASE SType
					CASE "D" '// 날짜별 이벤트 통계
						call oReport.GetExhibitionStatisticsByDateDataMart
				%>
				<tr bgcolor="#DDDDFF">
					<td width="90" align="center">구매일</td>
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
					<td width="500" align="center">그래프</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fselldate %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= oReport.ExhibitionReportList(i).Fsellcnt %></td>
					<td width="500" style="text-align:left;">
						<%
							'그래프 길이 계산 (2008.07.08;허진원 수정)
							if oReport.maxc>0 then
								grpWidth = Clng(oReport.ExhibitionReportList(i).Fselltotal/oReport.maxc*400)
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
					CASE "T"  '// 상품별 이벤트 통계
						call oReport.GetExhibitionStatisticsByItemDataMart
				%>
				<tr bgcolor="#EDEDFF">
					<td width="150" align="center" rowspan="2">브랜드</td>
					<td width="90" align="center" rowspan="2">아이템번호</td>
					<td rowspan="2">이미지</td>
					<td width="70" align="center" colspan="2">계</td>
					<td width="70" align="center" colspan="2">PC웹</td>
					<td width="70" align="center" colspan="2">모바일웹</td>
					<td width="70" align="center" colspan="2">APP</td>
					<td width="70" align="center" rowspan="2">Wish</td>
				</tr>
				<tr bgcolor="#EDEDFF">
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fmakerid %></td>
					<td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oReport.ExhibitionReportList(i).FItemid %>" target="_blank" title="미리보기"><%= oReport.ExhibitionReportList(i).FItemid %></a></td>
					<td><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(oReport.ExhibitionReportList(i).FItemid)%>/<%=oReport.ExhibitionReportList(i).Fsmallimage%>"></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_PC,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_PC,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_mobile,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_mobile,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellsum_App,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fsellcnt_App,0) %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).FwishCnt,0) %>건</td>
				</tr>
				<% next %>
				<% end if %>
				<%
					CASE "M"  '// 브랜드별 이벤트 통계
						call oReport.GetExhibitionStatisticsByMakerIDDataMart
				%>
				<tr bgcolor="#DDDDFF">
					<td width="150" align="center">브랜드</td>
					<td width="70" align="center">판매액</td>
					<td width="70" align="center">판매갯수</td>
					<td width="500" align="center">그래프</td>
				</tr>
				<% if oReport.FResultCount > 0 then %>
				<% for i=0 to oReport.FResultCount-1 %>
				<%
				t_TotalCost = t_TotalCost + oReport.ExhibitionReportList(i).Fselltotal
				t_FTotalNo  = t_FTotalNo + oReport.ExhibitionReportList(i).Fsellcnt
				%>
				<tr bgcolor="#FFFFFF">
					<td align="center"><%= oReport.ExhibitionReportList(i).Fmakerid %></td>
					<td align="right"><%= FormatNumber(oReport.ExhibitionReportList(i).Fselltotal,0) %></td>
					<td align="right"><%= oReport.ExhibitionReportList(i).Fsellcnt %></td>
					<td style="text-align:left;">
						<%
							'그래프 길이 계산 (2008.07.08;허진원 수정)
							if oReport.maxc>0 then
								grpWidth = Clng(oReport.ExhibitionReportList(i).Fselltotal/oReport.maxc*400)
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
						response.write "오류발생,다시 시도"
					END SELECT
				%>
			</table>
			<div>
				<table class="tbType1 listTb">
					<tr>
						<td> 총합금액 <%= FormatNumber(t_TotalCost,0) %> / 갯수 <%= FormatNumber(t_FTotalNo,0) %></td>
					</tr>
				</table>
			</div>
 		</div>
	</div>
</div>
<%
set oReport = Nothing
%>
<iframe src="about:blank" name="iframeDB1" width="0" height="0">
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
