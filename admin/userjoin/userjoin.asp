<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 회원가입 현황
' History : 2008.01.29 한용민 개발
'           2009.02.10 허진원 수정; 날짜함수 적용
'           2011.07.14 허진원 수정; 20, 30대 전후반으로 구분
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/userjoin/userjoincls.asp"-->

<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,defaultdate1 ,gubun,i ,joinSex , joinAreaSido,joinPath
dim ouserjoinlist ,ouserjoinsexcount ,ouserjoinarealist, ouserjoinchannellist, ouserjoin_areacount, ouserjoin_channelcount
dim strJson, strTemp1,strTemp2 ,joinAreaSidocount_total, joinPathcount_total
dim j, currDate, arrSize, page
	page = requestcheckvar(request("page"),10)
	gubun = request("gubun")
	if gubun = "" then gubun = "sex"
	joinSex = request("joinSex")
	joinAreaSido = request("joinAreaSido")
	joinPath = request("joinPath")
	defaultdate1 = dateadd("d",-7,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 7이전까지 검색
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = year(defaultdate1)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = month(defaultdate1)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = day(defaultdate1)
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then
		mm2 = month(now)
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)
	menupos = request("menupos")

if page = "" then page = 1

'연령별
if gubun = "sex" then
	set ouserjoinlist = new cuserjoinlist
	ouserjoinlist.frectjoinAreaSido = joinAreaSido
	ouserjoinlist.frectjoinPath = joinPath
	ouserjoinlist.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoinlist.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoinlist.fuserjoinlist()

	set ouserjoinsexcount = new cuserjoinlist
	ouserjoinsexcount.frectjoinAreaSido = joinAreaSido
	ouserjoinsexcount.frectjoinPath = joinPath
	ouserjoinsexcount.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoinsexcount.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoinsexcount.fuserjoin_sex()

' 가입경로
elseif (gubun = "channel") then
	set ouserjoinchannellist = new cuserjoinlist
	ouserjoinchannellist.FPageSize = 500
	ouserjoinchannellist.FCurrPage = page
	ouserjoinchannellist.frectjoinSex = joinSex
	ouserjoinchannellist.frectjoinPath = joinPath
	ouserjoinchannellist.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoinchannellist.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoinchannellist.fuserjoinchannellist()

	set ouserjoin_channelcount = new cuserjoinlist
	ouserjoin_channelcount.frectjoinSex = joinSex
	ouserjoin_channelcount.frectjoinPath = joinPath
	ouserjoin_channelcount.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoin_channelcount.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoin_channelcount.fuserjoin_channel()

	arrSize = DateDiff("d", ouserjoinchannellist.FRectStartdate, ouserjoinchannellist.FRectEndDate) + 1
	if (arrSize <= 0) then
		arrSize = 1
	elseif (arrSize > 7) then
		arrSize = 7
	end if

' 지역별
else
	set ouserjoinarealist = new cuserjoinlist
	ouserjoinarealist.frectjoinSex = joinSex
	ouserjoinarealist.frectjoinPath = joinPath
	ouserjoinarealist.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoinarealist.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoinarealist.fuserjoinarealist()

	set ouserjoin_areacount = new cuserjoinlist
	ouserjoin_areacount.frectjoinSex = joinSex
	ouserjoin_areacount.frectjoinPath = joinPath
	ouserjoin_areacount.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserjoin_areacount.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserjoin_areacount.fuserjoin_area()
end if

'그래프
if gubun = "sex" then		'연령별
	if ouserjoinsexcount.ftotalcount>0 then
		strJson = "["
		for i=0 to ouserjoinsexcount.ftotalcount -1
			strJson = strJson & "{"
			strJson = strJson & """label"":""" & ouserjoinsexcount.FItemList(i).fjoinSex & ""","
			strJson = strJson & """value"":""" & ouserjoinsexcount.FItemList(i).fjoinsexcount & """"
			strJson = strJson & "},"
		next
		strJson = strJson & "]"
	end if

elseif (gubun = "channel") then

	if ouserjoin_channelcount.ftotalcount>0 then
		strJson = "["
		for i=0 to ouserjoin_channelcount.ftotalcount -1
			strJson = strJson & "{"
			strJson = strJson & """label"":""" & ouserjoin_channelcount.FItemList(i).fjoinPath & ""","
			strJson = strJson & """value"":""" & ouserjoin_channelcount.FItemList(i).fjoinPathcount & """"
			strJson = strJson & "},"
		next
		strJson = strJson & "]"
	end if

else				'지역별

	if ouserjoin_areacount.ftotalcount>0 then
		strJson = "["
		for i=0 to ouserjoin_areacount.ftotalcount -1
			strJson = strJson & "{"
			strJson = strJson & """label"":""" & ouserjoin_areacount.FItemList(i).fjoinAreaSido & ""","
			strJson = strJson & """value"":""" & ouserjoin_areacount.FItemList(i).fjoinAreaSidocount & """"
			strJson = strJson & "},"
		next
		strJson = strJson & "]"
	end if
end if
%>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script type="text/javascript">

function frmsubmit(page){
	<% if (gubun = "channel") then %>
		//일단위
		var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
		var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
		var diffDay = 0;
		var start_yyyy = startdate.substring(0,4);
		var start_mm = startdate.substring(5,7);
		var start_dd = startdate.substring(8,startdate.length);
		var sDate = new Date(start_yyyy, start_mm-1, start_dd);
		var end_yyyy = enddate.substring(0,4);
		var end_mm = enddate.substring(5,7);
		var end_dd = enddate.substring(8,enddate.length);
		var eDate = new Date(end_yyyy, end_mm-1, end_dd);

		diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

		if (diffDay>92){
			alert("3개월까지 검색이 가능 합니다.");
			return;
		}
	<% end if %>

	frm.page.value=page;
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="radio" name="gubun" value="sex" <% if gubun = "sex" then response.write "checked" %>> 연령별
		<input type="radio" name="gubun" value="area" <% if gubun = "area" then response.write "checked" %>> 지역별
		<input type="radio" name="gubun" value="channel" <% if gubun = "channel" then response.write "checked" %>> 가입경로
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<% if gubun = "sex" then %>
	<% if ouserjoinlist.ftotalcount > 0 then %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td>
			<div id="chartdiv1" align="center"></div>
			<script type="text/javascript">
				var chart = new FusionCharts({
					type: 'doughnut2d',
					renderAt: 'chartdiv1',
					width: '100%',
					height: '400',
					dataFormat: 'json',
					dataSource: {
						"chart":{
							"caption": "총 가입수",
							"subCaption": "성별 비율(%)",
							"xAxisName": "성별",
							"yAxisName": "가입수",
							"numberSuffix": "명",
							"theme": "fusion",
							"formatNumberScale":"0",         // 천단위자동 변환 여부; 0:안함, 1:자동변환
							"formatNumber":"1"               // 천단위 쉼표 표시여부
						},
						"data" : <%=strJson%>
					}
				}).render();
			</script>
			</td>
		</tr>
		</table>
		<br>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td align="center">세부검색</td>
			<td align="left" colspan="5">&nbsp;
				지역 : <% DrawjoinAreaSido "joinAreaSido",joinAreaSido %>
				가입경로 : <% DrawjoinPath "joinPath",joinPath %>
				<a href="javascript:frmsubmit('1');"><image src="/admin/images/search2.gif" border="0"></a>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >연령</td>
			<td >전체</td>
			<td >남성</td>
			<td >여성</td>
		</tr>
		<% for  i = 0 to ouserjoinlist.ftotalcount -1 %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >총계</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_total_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_total_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >0-9세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_0_9_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_0_9_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_0_9_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >10-14세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_10_14_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_10_14_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_10_14_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >15-19세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_15_19_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_15_19_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_15_19_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >20-24세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_20_24_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_20_24_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_20_24_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >25-29세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_25_29_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_25_29_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_25_29_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >30-34세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_30_34_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_30_34_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_30_34_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >35-39세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_35_39_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_35_39_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_35_39_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >40-44세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_40_44_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_40_44_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_40_44_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >45-49세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_45_49_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_45_49_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_45_49_count %></td>
			</td>
		</tr>
		<!--
		<tr bgcolor="#FFFFFF" align="center">
			<td >50-59세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_50_59_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_50_59_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_50_59_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >60-69세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_60_69_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_60_69_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_60_69_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >70-79세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_70_79_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_70_79_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_70_79_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >80-89세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_80_89_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_80_89_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_80_89_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >90-99세</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_90_99_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_90_99_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_90_99_count %></td>
			</td>
		</tr>
		-->
		<tr bgcolor="#FFFFFF" align="center">
			<td >50세이상</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_50_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_50_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_50_count %></td>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" align="center">
			<td >미지정</td>
			<td ><%= ouserjoinlist.FItemList(i).ftotal_etc_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexman_etc_count %></td>
			<td ><%= ouserjoinlist.FItemList(i).fsexgirl_etc_count %></td>
			</td>
		</tr>
		<% next %>
		</table>

	<% else %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td>검색 결과가 없습니다.</td>
		</tr>
		</table>
	<% end if %>

<% elseif (gubun = "channel") then %>
	<% if ouserjoinchannellist.ftotalcount > 0 then %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF" align="center">
				<td>
					<div id="chartdiv2" align="center"></div>
					<script type="text/javascript">
						var chart = new FusionCharts({
							type: 'column2d',
							renderAt: 'chartdiv2',
							width: '100%',
							height: '400',
							dataFormat: 'json',
							dataSource: {
								"chart":{
									"caption": "총 가입수",
									"subCaption": "가입경로(명)",
									"xAxisName": "가입경로",
									"yAxisName": "가입수",
									"numberSuffix": "명",
									"theme": "fusion",
									"formatNumberScale":"0",         // 천단위자동 변환 여부; 0:안함, 1:자동변환
									"formatNumber":"1"               // 천단위 쉼표 표시여부
								},
								"data" : <%=strJson%>
							}
						}).render();
					</script>
				</td>
			</tr>
		</table>
		<br>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
		<tr bgcolor="#FFFFFF">
			<td align="center">세부검색</td>
			<td align="left" colspan="2">
				* 성별 : <% DrawjoinSex "joinSex",joinSex %>
				&nbsp;&nbsp;
				* 가입경로 : <% DrawjoinPath "joinPath",joinPath %>
				<input type="button" class="button_s" value="검색" onClick="frmsubmit('1');">
			</td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="3">
				검색결과 : <b><%= ouserjoinchannellist.FTotalCount %></b>
				&nbsp;
				페이지 : <b><%= page %>/ <%= ouserjoinchannellist.FTotalPage %></b>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>가입일</td>
			<td>가입경로</td>
			<td>가입수</td>
		</tr>
		<%
		for i = 0 to ouserjoinchannellist.FResultCount -1
		joinPathcount_total = joinPathcount_total + Clng(ouserjoinchannellist.FItemList(i).fjoinPath_count)
		%>
		<tr bgcolor="#FFFFFF" align="center">
			<td ><%= ouserjoinchannellist.FItemList(i).fjoindate %></td>
			<td><%= ouserjoinchannellist.FItemList(i).fjoinPath %></td>
			<td >
				<%= ouserjoinchannellist.FItemList(i).fjoinPath_count %>
			</td>
		</tr>
		<% next %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="2">합계</td>
			<td ><%= joinPathcount_total %></td>
		</tr>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="3" align="center">
				<% if ouserjoinchannellist.HasPreScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= ouserjoinchannellist.StartScrollPage-1 %>')">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + ouserjoinchannellist.StartScrollPage to ouserjoinchannellist.StartScrollPage + ouserjoinchannellist.FScrollCount - 1 %>
				<% if (i > ouserjoinchannellist.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ouserjoinchannellist.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="javascript:frmsubmit('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
				<% next %>
				<% if ouserjoinchannellist.HasNextScroll then %>
				<span class="list_link"><a href="javascript:frmsubmit('<%= i %>')">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
		</table>
		<% else %>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF" align="center">
				<td>검색 결과가 없습니다.</td>
			</tr>
			</table>
		<% end if %>

<% else %>
	<% if ouserjoinarealist.ftotalcount > 0 then %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td>
				<div id="chartdiv2" align="center"></div>
				<script type="text/javascript">
					var chart = new FusionCharts({
						type: 'column2d',
						renderAt: 'chartdiv2',
						width: '100%',
						height: '400',
						dataFormat: 'json',
						dataSource: {
							"chart":{
								"caption": "총 가입수",
								"subCaption": "지역별(명)",
								"xAxisName": "가입지역",
								"yAxisName": "가입수",
								"numberSuffix": "명",
								"theme": "fusion",
								"formatNumberScale":"0",         // 천단위자동 변환 여부; 0:안함, 1:자동변환
								"formatNumber":"1"               // 천단위 쉼표 표시여부
							},
							"data" : <%=strJson%>
						}
					}).render();
				</script>
			</td>
		</tr>
		</table>
		<br>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td align="center">세부검색</td>
			<td align="left" colspan="5">&nbsp;
				성별 : <% DrawjoinSex "joinSex",joinSex %>
				가입경로 : <% DrawjoinPath "joinPath",joinPath %>
				<a href="javascript:frmsubmit('1');"><image src="/admin/images/search2.gif" border="0"></a>
			</td>
		</tr>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >지역</td>
			<td >성별</td>
			<td>가입수</td>
		</tr>
		<% for i = 0 to ouserjoinarealist.ftotalcount -1 %>
			<tr bgcolor="#FFFFFF" align="center">
				<td ><%= ouserjoinarealist.FItemList(i).fjoinAreaSido %></td>
				<td ><%= ouserjoinarealist.FItemList(i).fjoinSex %></td>
				<td >
					<%= ouserjoinarealist.FItemList(i).fjoinAreaSidocount %></td>
					<% joinAreaSidocount_total = joinAreaSidocount_total + Clng(ouserjoinarealist.FItemList(i).fjoinAreaSidocount) %>
				</td>
			</tr>
		<% next %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td colspan="2">합계</td>
			<td ><%= joinAreaSidocount_total %></td>
		</tr>
		</table>
	<% else %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td>검색 결과가 없습니다.</td>
		</tr>
		</table>
	<% end if %>
<% end if %>

</form>

<%
if gubun = "sex" then
	set ouserjoinlist = nothing
	set ouserjoinsexcount = nothing

'지역별
else
	set ouserjoinarealist = nothing
	set ouserjoin_areacount = nothing
end if

'####################### 연령별검색란 ########################	
Sub DrawjoinSex(boxname, joinSex)
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"
		if joinSex ="" then
			response.write "selected"
		end if
	response.write ">전체</option>"

	response.write "<option value='남'"
		if joinSex ="남" then
			response.write "selected"
		end if
	response.write ">남자</option>"

	response.write "<option value='여'"
		if joinSex ="여" then
			response.write "selected"
		end if
	response.write ">여자</option>"
	response.write "</select>"
End Sub
'######################### 지역검색란 ##########################	
Sub DrawjoinAreaSido(boxname, joinAreaSido)
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"
	response.write "<option value=''"
		if joinAreaSido ="" then
			response.write "selected"
		end if
	response.write ">전체</option>"

	'사용자 검색 옵션 내용 DB에서 가져오기
		userquery = "select distinct joinAreaSido"
		userquery = userquery + " from db_datamart.dbo.tbl_user_join_log"
		userquery = userquery + " where joinAreaSido <>''"
		userquery = userquery + " group by joinAreaSido"
	db3_rsget.Open userquery, db3_dbget, 1

	if not db3_rsget.EOF then
		do until db3_rsget.EOF
			if Lcase(joinAreaSido) = Lcase(db3_rsget("joinAreaSido")) then
				tem_str = " selected"
			end if
			response.write "<option value='" & db3_rsget("joinAreaSido") & "' " & tem_str & ">" & db2html(db3_rsget("joinAreaSido")) & "</option>"
			tem_str = ""
			db3_rsget.movenext
		loop
	end if
	db3_rsget.close
	response.write "</select>"
End Sub
'######################## 가입경로 #########################	
Sub DrawjoinPath(boxname, joinPath)
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"
	response.write "<option value=''"
		if joinPath ="" then
			response.write "selected"
		end if
	response.write ">전체</option>"

	'사용자 검색 옵션 내용 DB에서 가져오기
		userquery = "select (case when joinPath='' then '10X10' else joinPath end) as joinPath"
		userquery = userquery + " from db_datamart.dbo.tbl_user_join_log"
		userquery = userquery + " group by (case when joinPath='' then '10X10' else joinPath end)"
	db3_rsget.Open userquery, db3_dbget, 1

	if not db3_rsget.EOF then
		do until db3_rsget.EOF
			if Lcase(joinPath) = Lcase(db3_rsget("joinPath")) then
				tem_str = " selected"
			end if
			response.write "<option value='" & db3_rsget("joinPath") & "' " & tem_str & ">" & db2html(db3_rsget("joinPath")) & "</option>"
			tem_str = ""
			db3_rsget.movenext
		loop
	end if
	db3_rsget.close
	response.write "</select>"
End Sub
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
