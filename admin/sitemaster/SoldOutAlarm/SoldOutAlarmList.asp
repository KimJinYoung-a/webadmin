<% Option Explicit %>
<%
'###########################################################
' Description : 품절상품 입고알림 통계페이지
' Hieditor : 2018.02.27 원승현 생성
'			 2020.03.20 한용민 수정(테스트서버 환경셋팅, 회원등급 개편등급으로 변경, 장바구니건수 추가)
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/SoldOutAlarm/SoldOutAlarmcls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->

<%
	'// 기존에 대충 만들었던 top100 신청 리스트 주소 : /admin/report/SoldOutPushItemTop100.asp
	Dim startDate '// 기간 시작일
	Dim endDate '// 기간 종료일
	Dim cateCode '// 카테고리 코드
	Dim makerId '// 브랜드(메이커아이디)
	Dim searchKeyword '// 검색값
	Dim searchGubun '// (1-상품코드, 2-상품명)
	Dim oSoldOutAlarmList, RecGraphSoldOutAlarm, ColGraphSoldOutAlarm
	Dim loginUserId
	Dim currpage, research, pagesize, i
	Dim cdl, cdm, cds, dispCate, searchEndDate
	Dim RectGraphData, ColGraphData

	loginUserId = session("ssBctId")
	currpage = requestcheckvar(request("page"), 20)
	startDate = requestcheckvar(request("startDate"), 20)
	endDate = requestcheckvar(request("endDate"), 20)
	research = requestcheckvar(request("research"), 20)
	searchKeyword = requestcheckvar(request("searchKeyword"), 500)
	searchGubun = requestcheckvar(request("searchGubun"), 1)
	cateCode = requestcheckvar(request("disp"), 50)
	makerId = requestcheckvar(request("makerId"), 100)
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	dispCate = requestCheckvar(request("disp"),16)

	If Trim(startDate) = "" Then
		startDate = Left(Now(), 7)&"-01"
	End If

	If Trim(endDate) = "" Then
		endDate = DateAdd("m", 1, startdate)
		searchEndDate = endDate
	Else
		searchEndDate = DateAdd("d", 1, endDate)
	End If

	If Trim(currpage)="" Then
		currpage = "1"
	End If
	pagesize = 50

	'// 리스트를 가져온다.
	set oSoldOutAlarmList = new CGetSoldOutAlarm
		oSoldOutAlarmList.FRectcurrpage = currpage
		oSoldOutAlarmList.FRectpagesize = pagesize
		oSoldOutAlarmList.FRectStartDate = startDate
		oSoldOutAlarmList.FRectEndDate = searchEndDate
		If Trim(research)="on" Then
			oSoldOutAlarmList.FRectSearchGubun = searchGubun
			oSoldOutAlarmList.FRectSearchKeyword = searchKeyword
			oSoldOutAlarmList.FRectCateCode = dispCate
			oSoldOutAlarmList.FRectMakerId = makerId
		End If
		oSoldOutAlarmList.GetSoldOutAlarmList()

	'// 원형 그래프 데이터를 가져온다.
	Set RecGraphSoldOutAlarm = new CGetSoldOutAlarm
		RecGraphSoldOutAlarm.FRectStartDate = startDate
		RecGraphSoldOutAlarm.FRectEndDate = searchEndDate
		If Trim(research)="on" Then
			RecGraphSoldOutAlarm.FRectSearchGubun = searchGubun
			RecGraphSoldOutAlarm.FRectSearchKeyword = searchKeyword
			RecGraphSoldOutAlarm.FRectCateCode = dispCate
			RecGraphSoldOutAlarm.FRectMakerId = makerId
		End If
		RecGraphSoldOutAlarm.GetUserLevelAlarm()

	'// 카테고리별 신청 그래프 데이터를 가져온다.
	Set ColGraphSoldOutAlarm = new CGetSoldOutAlarm
		ColGraphSoldOutAlarm.FRectStartDate = startDate
		ColGraphSoldOutAlarm.FRectEndDate = searchEndDate
		If Trim(research)="on" Then
			ColGraphSoldOutAlarm.FRectSearchGubun = searchGubun
			ColGraphSoldOutAlarm.FRectSearchKeyword = searchKeyword
			ColGraphSoldOutAlarm.FRectCateCode = dispCate
			ColGraphSoldOutAlarm.FRectMakerId = makerId
		End If
		ColGraphSoldOutAlarm.GetCategoryAlarm()

	'// 원형 그래프 데이터
	if RecGraphSoldOutAlarm.FResultCount > 0 Then
		RectGraphData = ""
		for i = 0 to RecGraphSoldOutAlarm.FResultCount -1 
			RectGraphData = RectGraphData & "{'label': '"&RecGraphSoldOutAlarm.FUserLevelAlarmList(i).FGraphUserLevel&"', 'value': '"&RecGraphSoldOutAlarm.FUserLevelAlarmList(i).FGraphUserCount&"'},"
		Next
		RectGraphData = Left(RectGraphData, Len(RectGraphData) - 1)
	End If

	'// 카테고리 그래프 데이터
	if ColGraphSoldOutAlarm.FResultCount > 0 Then
		ColGraphData = ""
		for i = 0 to ColGraphSoldOutAlarm.FResultCount -1 
			ColGraphData = ColGraphData & "{'label': '"&ColGraphSoldOutAlarm.FCategoryAlarmList(i).FCateName1&"', 'value': '"&ColGraphSoldOutAlarm.FCategoryAlarmList(i).FCategoryCnt&"'},"
		Next
		ColGraphData = Left(ColGraphData, Len(ColGraphData) - 1)
	End If

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>
document.domain = "10x10.co.kr";

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "즐겨찾기에서 제외하시겠습니까?";
	} else {
		msg = "즐겨찾기에 추가하시겠습니까?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}


function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&startDate=<%=startDate%>&endDate=<%=endDate%>&searchKeyword=<%=searchKeyword%>&searchGubun=<%=searchGubun%>&disp=<%=dispCate%>&makerId=<%=makerId%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchSoldOutAlarm()
{
//	if ($("#deal").val()=="0"&&$("#open").val()=="A"&&$("#schWord").val()=="")
//	{
//		alert("검색을 하기 위해선 구분, 노출여부, 키워드검색 중 하나를\n선택해주시거나 입력해주셔야 합니다.");
//		return;
//	}
	document.frm1.submit();
}

</script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<script type="text/javascript">
FusionCharts.ready(function(){
	var myChart = new FusionCharts({
		"type": "pie2d",
		"renderAt": "chartContainer",
		"width": "50%",
		"height": "300",
		"dataFormat": "json",
		"dataSource":  {
			"chart": {
				"caption": "기간 내 회원 등급별 신청 그래프",
				"subcaption": "",
				"showvalues": "1",
				"showpercentvalues": "1",
				"showpercentintooltip": "0",
				"bgcolor": "#FFFFFF",
				"basefontcolor": "#400D1B",
				"showshadow": "0",
				"animation": "0",
				"showborder": "0",
				"divLineDashed": "1",
				"palettecolors": "#BE3243,#986667,#BE6F71,#CB999A,#DFC0B1,#E0D0D0"
			},
			"data": [
				<%=RectGraphData%>
			]
		}
	});
	myChart.render();
})

FusionCharts.ready(function(){
	var myChart2 = new FusionCharts({
		"type": "column2D",
		"renderAt": "chartContainer2",
		"width": "50%",
		"height": "300",
		"dataFormat": "json",
		"dataSource":  {
			"chart": {
				"caption": "기간 내 카테고리별 신청 그래프",
				"subCaption": "",
				"xAxisName": "",
				"yAxisName": "",
				"paletteColors": "#0075c2",
				"bgColor": "#ffffff",
				"borderAlpha": "20",
				"canvasBorderAlpha": "0",
				"usePlotGradientColor": "0",
				"plotBorderAlpha": "10",
				"placevaluesInside": "0",
				"rotatevalues": "0",
				"valueFontColor": "#ffffff",
				"showXAxisLine": "1",
				"xAxisLineColor": "#999999",
				"divlineColor": "#999999",
				"divLineDashed": "1",
				"showAlternateHGridColor": "0",
				"subcaptionFontBold": "0",
				"subcaptionFontSize": "10",
				"showValue":"1",
				"formatNumber":"1",
				"showLimits":"0",
				"placeValuesInside":"1"
			},
			"data": [
				<%=ColGraphData%>
			]
		}
	});
	myChart2.render();
})
</script>
<div class="">
	<%' 상단 검색폼 시작 %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/SoldOutAlarm/SoldOutAlarmList.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<p class="formTit">기간</p>
					<input type="text" id="startDate" name="startDate" value="<%=startDate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startDate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "startDate", trigger    : "startDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
					~
					<input type="text" id="endDate" name="endDate" value="<%=endDate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="endDate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "endDate", trigger    : "endDate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
				</li>
				<li>
					<label class="formTit">카테고리 :</label>
					<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
				</li>
				<li>
					<p class="formTit">브랜드 : </p>
					<input type="text" id="makerId" name="makerId" value="<%=makerId%>" class="formTxt" size="10" style="margin-bottom:13px;" readonly onclick="window.open('/admin/member/popBrandSearch.asp?frmName=frm1&compName=makerId&isjsdomain=o','popBrandSearch','width=800 height=400 scrollbars=yes resizable=yes');return false;"/>
				</li>
			</ul>
		</div>
		<dfn class="line"></dfn>
		<div class="search">
			<ul>
				<li>
					<label class="formTit" for="schWord">검색어 :</label>
					<select class="formSlt" id="searchGubun" name="searchGubun" title="검색구분">
						<option value="" <% If searchGubun = "" Then %> selected <% End If %>>선택해주세요.</option>
						<option value="1" <% If searchGubun = "1" Then %> selected <% End If %>>상품코드</option>
						<option value="2" <% If searchGubun = "2" Then %>selected<% End If %>>상품명</option>
					</select>
					<input type="text" class="formTxt" id="searchKeyword" name="searchKeyword" style="width:400px" placeholder="키워드를 입력하여 검색하세요." value="<%=searchKeyword%>" />
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="검색" onclick="goSearchSoldOutAlarm();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<span id="chartContainer"></span><span id="chartContainer2"></span>
		</div>
	</div>
	<div class="cont" style="overflow:inherit;">
		<div class="pad20">
			<div class="overHidden">
				<div class="ftLt">
					<% If Trim(research)="on" Then %>
						<input type="button" class="btnRegist btn bold fs12" value="검색초기화" onclick="document.location.href='/admin/sitemaster/SoldOutAlarm/SoldOutAlarmList.asp';" />
					<% End If %>
				</div>
			</div>
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">기간 내 총 신청상품수 : <strong><%=FormatNumber(oSoldOutAlarmList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList" style="position:sticky; top:0; height:33px;">
						<li>
							<p style="width:100px">1차 카테고리</p>
							<p style="width:100px">2차 카테고리</p>
							<p style="width:120px">브랜드</p>
							<p style="width:80px">상품코드</p>
							<p style="width:300px">상품명</p>
							<p style="width:80px">전체(건)</p>
							<p style="width:80px">PC(건)</p>
							<p style="width:80px">M/W(건)</p>
							<p style="width:80px">APP(건)</p>
							<!--
							<p style="width:80px">구매(건)</p>
							<p style="width:80px">장바구니(건)</p>
							-->
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<%'// 리스트를 가져온다. %>
						<% If oSoldOutAlarmList.FResultcount > 0 Then %>
							<li style="background-color:#C6C1C1">
								<p style="width:700px"><strong>기간 전체</strong></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FAlarmCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FPcCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FMobileCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FAppCount, 0)%></p>
							<!--
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FBuyCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.fbagunicnt, 0)%></p>
							-->
							</li>

							<% For i=0 To oSoldOutAlarmList.Fresultcount-1 %>
							<li>
								<p style="width:100px"><%=oSoldOutAlarmList.FSoldOutAlarmList(i).FCateName1%></p>
								<p style="width:100px"><%=oSoldOutAlarmList.FSoldOutAlarmList(i).FCateName2%></p>
								<p style="width:120px"><%=oSoldOutAlarmList.FSoldOutAlarmList(i).FMakerId%></p>
								<p style="width:80px"><a href="http://www.10x10.co.kr/<%=oSoldOutAlarmList.FSoldOutAlarmList(i).FItemId%>" target="_BLANK"><%=oSoldOutAlarmList.FSoldOutAlarmList(i).FItemId%></a></p>
								<p style="width:300px;text-align:left;"><%=oSoldOutAlarmList.FSoldOutAlarmList(i).FItemName%><% If oSoldOutAlarmList.FSoldOutAlarmList(i).FOptionCheck > 0 Then %>&nbsp;<span style="color:red;" onclick="window.open('/admin/sitemaster/SoldOutAlarm/PopSoldOutAlarmItemInfo.asp?startDate=<%=startDate%>&endDate=<%=endDate%>&itemId=<%=oSoldOutAlarmList.FSoldOutAlarmList(i).FItemId%>', 'optionDetail', 'width=1000, height=600');return false;">[옵션신청현황]</span><% End If %></p>								
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).FListTotalCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).FListPCCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).FListMobileCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).FListAppCount, 0)%></p>
							<!--
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).FListBuyCount, 0)%></p>
								<p style="width:80px"><%=FormatNumber(oSoldOutAlarmList.FSoldOutAlarmList(i).fbagunicnt, 0)%></p>
							-->
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%= fnDisplayPaging_New2017(currpage, oSoldOutAlarmList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="dimmed"></div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
</body>
</html>
<%
	Set oSoldOutAlarmList = Nothing
	Set RecGraphSoldOutAlarm = Nothing
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
