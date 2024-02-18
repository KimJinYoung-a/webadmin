<% Option Explicit %>
<%
'###########################################################
' Description : 품절상품 입고알림 팝업(상품상세)
' Hieditor : 2018.02.28 원승현 생성
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
	Dim startDate, endDate, itemId, i
	Dim ItemBasicInfoAlarm, ItemOptionInfoAlarm
	Dim totalCnt, pcTotalCnt, appTotalCnt, mobileTotalCnt, buyTotalCnt

	startDate = requestcheckvar(request("startDate"), 20)
	endDate = requestcheckvar(request("endDate"), 20)
	itemId = requestcheckvar(request("itemId"), 30)

	'// 상품 기본정보
	Set ItemBasicInfoAlarm = new CGetSoldOutAlarm
		ItemBasicInfoAlarm.FRectItemId = itemId
		ItemBasicInfoAlarm.GetItemBasicInfoAlarm()

	'// 옵션 상세정보
	Set ItemOptionInfoAlarm = new CGetSoldOutAlarm
		ItemOptionInfoAlarm.FRectItemId = itemId
		ItemOptionInfoAlarm.FRectStartDate = startDate
		ItemOptionInfoAlarm.FRectEndDate = endDate
		ItemOptionInfoAlarm.GetItemOptionInfoAlarm()

		totalCnt = 0
		pcTotalCnt = 0
		appTotalCnt = 0
		mobileTotalCnt = 0
		buyTotalCnt = 0
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
</script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>


<div class="">

	<div class="cont">
		<div class="pad20">
			<div class="pieceList">
				<div class="tbListWrap">
					<ul class="thDataList"><li><p>상품기본정보</p></li></ul>
					<ul class="thDataList">
						<li>
							<p style="width:120px">이미지</p>
							<p style="width:80px">브랜드</p>
							<p style="width:100px">상품코드</p>
							<p style="width:200px">상품명</p>
							<p style="width:80px">배송구분</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<li>
							<p style="width:120px"><img src="<%=ItemBasicInfoAlarm.FItemBasicInfoAlarm(0).FListImage%>"></p>
							<p style="width:80px"><%=ItemBasicInfoAlarm.FItemBasicInfoAlarm(0).FBrandName%></p>
							<p style="width:100px"><%=ItemBasicInfoAlarm.FItemBasicInfoAlarm(0).FItemId%></p>
							<p style="width:200px;"><%=ItemBasicInfoAlarm.FItemBasicInfoAlarm(0).FItemName%></p>								
							<p style="width:80px"><%=ItemBasicInfoAlarm.FItemBasicInfoAlarm(0).FBaesongGubun%></p>
						</li>
					</ul>
				</div>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
				<div class="tbListWrap">
					<ul class="thDataList"><li><p>옵션별 신청 현황</p></li></ul>
					<ul class="thDataList">
						<li>
							<p style="width:150px">옵션명</p>
							<p style="width:80px">전체(건)</p>
							<p style="width:80px">PC(건)</p>
							<p style="width:80px">M/W(건)</p>
							<p style="width:80px">APP(건)</p>
							<p style="width:80px">구매(건)</p>
						</li>
					</ul>
					<ul id="sortable" class=" tbDataList">
						<% if ItemOptionInfoAlarm.FResultCount > 0 Then %>
							<% for i = 0 to ItemOptionInfoAlarm.FResultCount -1 %>
								<li>
									<p style="width:150px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionName%></p>
									<p style="width:80px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionTotalCnt%></p>
									<p style="width:80px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionPcCnt%></p>
									<p style="width:80px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionMobileCnt%></p>
									<p style="width:80px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionAppCnt%></p>
									<p style="width:80px"><%=ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionBuyCnt%></p>
								</li>
								<%
									totalCnt = totalCnt + ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionTotalCnt
									pcTotalCnt = pcTotalCnt + ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionPcCnt
									mobileTotalCnt = mobileTotalCnt + ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionMobileCnt
									appTotalCnt = appTotalCnt + ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionAppCnt
									buyTotalCnt = buyTotalCnt + ItemOptionInfoAlarm.FItemOptionInfoAlarm(i).FOptionBuyCnt
								%>
							<% Next %>
						<% End If %>
						<li style="background-color:#C6C1C1">
							<p style="width:150px">전체</p>
							<p style="width:80px"><%=totalCnt%></p>
							<p style="width:80px"><%=pcTotalCnt%></p>
							<p style="width:80px"><%=mobileTotalCnt%></p>
							<p style="width:80px"><%=appTotalCnt%></p>
							<p style="width:80px"><%=buyTotalCnt%></p>
						</li>
					</ul>
				</div>

			</div>
		</div>
	</div>
</div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
</body>
</html>
<%
	Set ItemBasicInfoAlarm = Nothing
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
