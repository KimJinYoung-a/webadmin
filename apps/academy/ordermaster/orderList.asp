<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 주문 관리"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
dim statediv, sdiv, sortupdown, isFilter, searchdiv
dim page, searchtxt, MakerID, odiv, ix, iy, sdate, edate

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)

odiv = RequestCheckVar(request("odiv"),2)
searchdiv = RequestCheckVar(request("searchdiv"),1)
searchtxt = RequestCheckVar(request("searchtxt"),32)
statediv = RequestCheckVar(request("statediv"),1)
sortupdown = RequestCheckVar(request("sortupdown"),1)
sdiv = RequestCheckVar(request("sdiv"),10)
page = RequestCheckVar(request("page"),10)
sdate = RequestCheckVar(request("sdate"),10)
edate = RequestCheckVar(request("edate"),10)
isFilter = RequestCheckVar(request("isFilter"),1)

If (MakerID="") Then
	Response.Write "<script>alert('주문 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

if (odiv="") then odiv="S"
if (statediv="") then statediv="0"
if (page="") then page=1
If (sortupdown="") Then sortupdown="u"
If searchdiv="" Then searchdiv=0
'필터 파라메터
dim strParam
strParam = "statediv=" & statediv & "&sortupdown=" & sortupdown & "&sdiv=" & sdiv & "&sdate=" & sdate & "&edate=" & edate & "&searchdiv=" & searchdiv & "&searchtxt=" & searchtxt
'==================================================================================++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
dim ojumun
If odiv="C" Then
set ojumun = New CJumunMaster
ojumun.FPageSize = 12
ojumun.FCurrPage = page
ojumun.FRectMakerid=MakerID
ojumun.FRectStateDIV = statediv
ojumun.FRectSortUpDown = sortupdown
If searchdiv=1 Then
	ojumun.FRectOrderSerial=searchtxt
ElseIf searchdiv=3 Then
	ojumun.FRectUserName=searchtxt
End If
ojumun.FRectSearchType = "upcheview"

ojumun.GetCSASMasterList
Else
set ojumun = new CJumunMaster
ojumun.FPageSize = 12
ojumun.FCurrPage = page
ojumun.FRectDesignerID=MakerID
ojumun.FRectStateDIV = statediv
ojumun.FRectSortUpDown = sortupdown
ojumun.FRectOrderDiv = odiv
ojumun.FRectStartDate = sdate
ojumun.FRectEndDate = edate
ojumun.FRectSearchDIV = searchdiv
ojumun.FRectSearchTxt = searchtxt
ojumun.DesignerDateBaljuList
End If

'뱃지 카운트 체크
Dim StandByConfirmCnt, BMiBeasongCnt, OrderCSCnt, UpdateCheck
GetCheckIconBadgeCount MakerID, StandByConfirmCnt, BMiBeasongCnt, OrderCSCnt, UpdateCheck

dim i
%>
<script>
$(function() {
	// search button control
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});

	// search box hidden scroll top auto change
	var schH = $(".artSearchTop").outerHeight();
	var tabT = $(".listTab").offset().top;
	$("body").css("min-height",schH+$(window).height());
	setTimeout(function(){
		$('html, body').animate({scrollTop:schH-tabT}, 'fast');
	}, 300);
});

function jsGoPage(iP){
	document.searchForm.page.value = iP;
	document.searchForm.submit();
}
function fnGoTapPage(param){
	if(param=="S"){
		location.href="/apps/academy/ordermaster/OrderList.asp?odiv="+param;
	}else if(param=="D"){
		location.href="/apps/academy/ordermaster/OrderList.asp?odiv="+param;
	}else{
		location.href="/apps/academy/ordermaster/OrderList.asp?odiv="+param;
	}
}

function fnSearchFilterSet(callbackval){
	callbackval = Base64.decode(callbackval);
	//var catearr = callbackval.replace(/ /g, "");
	//var catearr2 = catearr.replace(/,/g, "','");
	//var catearr3=eval("['" + catearr2 + "']");
	var jsontxt = JSON.parse(callbackval);
	//alert(jsontxt.statediv);
	$("#statediv").val(jsontxt.statediv);
	$("#sdiv").val(jsontxt.sdiv);
	$("#sortupdown").val(jsontxt.ssort);
	$("#sdate").val(jsontxt.startdate);
	$("#edate").val(jsontxt.enddate);
	$("#isFilter").val(jsontxt.filter);
	document.searchForm.submit();
}

function fnSearchList(){
	if(document.searchForm.searchdiv.value==0){
		alert("구분을 선택해주세요.");
		document.searchForm.searchdiv.focus();
	}else if(document.searchForm.searchtxt.value==""){
		alert("검색어를 입력해주세요.");
		document.searchForm.searchtxt.focus();
	}else{
		document.searchForm.submit();
	}
}

function fnThisPageRelold(){
	window.location.reload(true);
}
</script>
</head>
<body onload="document.body.scrollTop=document.cookie!" onunload="document.cookie!=document.body.scrollTop">
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">주문관리</h1>
			<div class="orderManage">
				<ul class="listTab">
					<li<% If odiv="S" Then %> class="current"<% End If %> onclick="fnGoTapPage('S');"><div>주문접수<% If StandByConfirmCnt>0 Then %>(<%=StandByConfirmCnt%>)<% End If %></div></li>
					<li<% If odiv="D" Then %> class="current"<% End If %> onclick="fnGoTapPage('D');"><div>주문처리<% If BMiBeasongCnt>0 Then %>(<%=BMiBeasongCnt%>)<% End If %></div></li>
					<li<% If odiv="C" Then %> class="current"<% End If %> onclick="fnGoTapPage('C');"><div>관련CS<% If OrderCSCnt>0 Then %>(<%=OrderCSCnt%>)<% End If %></div></li>
				</ul>
				<div class="artListCont">
				<form name="searchForm" id="searchForm" method="get" style="margin:0px;" onSubmit="fnSearchList();return false;" action="">
				<input type="hidden" name="page" id="page" value="1">
				<input type="hidden" name="odiv" id="odiv" value="<%=odiv%>">
				<input type="hidden" name="statediv" id="statediv" value="<%=statediv%>">
				<input type="hidden" name="sdiv" id="sdiv" value="<%=sdiv%>">
				<input type="hidden" name="sortupdown" id="sortupdown" value="<%=sortupdown%>">
				<input type="hidden" name="sdate" id="sdate" value="<%=sdate%>">
				<input type="hidden" name="edate" id="edate" value="<%=edate%>">
				<input type="hidden" name="isFilter" id="isFilter" value="<%=isFilter%>">
					<% If isFilter="Y" Or searchtxt<>"" Or ojumun.FresultCount > 0 Then %>
					<div class="artSearchTop">
						<div class="searchInput hasOpt"><!-- for dev msg : 검색의 구분선택이 있는 경우 hasOpt 클래스 붙여주세요 -->
							<span class="schSlt">
								<select name="searchdiv">
									<% If odiv="C" Then %>
									<option value="0"<% If searchdiv=0 Then Response.write " selected"%>>구분선택</option>
									<option value="1"<% If searchdiv=1 Then Response.write " selected"%>>주문번호</option>
									<option value="3"<% If searchdiv=3 Then Response.write " selected"%>>구매자</option>
									<% Else %>
									<option value="0"<% If searchdiv=0 Then Response.write " selected"%>>구분선택</option>
									<option value="1"<% If searchdiv=1 Then Response.write " selected"%>>주문번호</option>
									<option value="2"<% If searchdiv=2 Then Response.write " selected"%>>작품코드</option>
									<option value="3"<% If searchdiv=3 Then Response.write " selected"%>>구매자</option>
									<option value="4"<% If searchdiv=4 Then Response.write " selected"%>>주문작품명</option>
									<% End If %>
								</select>
							</span>
							<input type="Search" name="searchtxt" placeholder="<% If odiv="C" Then %>주문번호, 구매자<% Else %>주문번호, 작품코드, 구매자, 주문작품명<% End If %> 검색" value="<%=searchtxt%>" onKeyPress="if (event.keyCode == 13){ fnSearchList(); return false;}" />
							<button type="button" class="btnSearch" onClick="fnSearchList();return false;">검색</button>
						</div>
						<div class="btnFilter <%=chkIIF(isFilter="Y","filterActive","")%>">
							<button type="button" onclick="fnAPPpopupSearchFilter('<%=g_AdminURL%>/apps/academy/ordermaster/popup/popFilter.asp?odiv=<%=odiv%>&<%=strParam%>')">필터</button>
						</div>
					</div>
					<% End If %>
<% If odiv<>"C" Then %>
				<% if ojumun.FresultCount > 0 then %>
<%
Dim OrderItemDic, SameOrderCheck, LoopCount, SameOrderCheck2, ItemCount, SameOrderSerial, OrderCanCelCnt, BadgeCount
Dim BeasongCnt, BeasongState, MibeasongCnt, MiChulGoCheck, BeasongStateName, BeasongStateClass, TotalOrderCnt
BeasongCnt=0
MibeasongCnt=0
BadgeCount=0
Set OrderItemDic = Server.CreateObject("Scripting.Dictionary")
For i=0 To ojumun.FresultCount-1

ReDim OrderInfoArr(ojumun.FresultCount-1)
OrderInfoArr(i) = Array(ojumun.FMasterItemList(i).FIpkumdate,ojumun.FMasterItemList(i).FOrderserial,ojumun.FMasterItemList(i).IpkumDivName,ojumun.FMasterItemList(i).FListimage,ojumun.FMasterItemList(i).Fitemname,ojumun.FMasterItemList(i).Fsongjangno,ojumun.FMasterItemList(i).Fcode,ojumun.FMasterItemList(i).Frequiremakeday,ojumun.FMasterItemList(i).Fupcheconfirmdate,ojumun.FMasterItemList(i).FCancelYn,ojumun.FMasterItemList(i).FMCancelYn)
OrderItemDic.Add ojumun.FMasterItemList(i).Fidx, OrderInfoArr(i)
Next

SameOrderCheck=""
SameOrderCheck2=""
SameOrderSerial=""
BeasongStateName=""
BeasongStateClass=""
LoopCount=0
ItemCount=0
OrderCanCelCnt=0
%>
					<div class="orderListWrap">
						<ul class="artList">
							<% For Each ix In OrderItemDic %>
							<% For Each iy In OrderItemDic %>
							<%
							If OrderItemDic.item(ix)(1)=OrderItemDic.item(iy)(1) And SameOrderCheck<>OrderItemDic.item(iy)(1) Then
								SameOrderCheck=OrderItemDic.item(iy)(1)
								SameOrderCheck2="Y"
							End If
							If OrderItemDic.item(ix)(1)=OrderItemDic.item(iy)(1) Then
								 SameOrderSerial=OrderItemDic.item(iy)(1)
							End If
							If SameOrderSerial=OrderItemDic.item(iy)(1) Then
								ItemCount = ItemCount + 1
								'일부 출고 확인
								If OrderItemDic.item(iy)(5) <> "" Then
									BeasongCnt=BeasongCnt+1
								End If
								If OrderItemDic.item(iy)(6) <> "" Then
									MibeasongCnt=MibeasongCnt+1
								End If
								If (DateDiff("d",DateAdd("d",OrderItemDic.item(iy)(7)+2,OrderItemDic.item(iy)(8)),now())>0) Then
									MiChulGoCheck=MiChulGoCheck+1
								End If
								If OrderItemDic.item(iy)(9)="Y" Or OrderItemDic.item(iy)(10)="Y" Then
									OrderCanCelCnt=OrderCanCelCnt+1
								End If

								'Response.write SameOrderSerial & " / " & ItemCount & "<br>"
								If ItemCount = BeasongCnt Then
									BeasongState="0"
									BeasongStateName="출고완료"
									BeasongStateClass="releaseFin"
								ElseIf ItemCount > BeasongCnt And BeasongCnt>0 Then
									BeasongState="1"
									BeasongStateName="일부출고"
									BeasongStateClass="releaseIng"
								ElseIf MibeasongCnt>0 Or MiChulGoCheck>0 Then
									BeasongState="2"
									BeasongStateName="미출고"
									BeasongStateClass="undeliver"
								ElseIf ItemCount=OrderCanCelCnt Then
									BeasongState="4"
									BeasongStateName="주문취소"
									BeasongStateClass="odrCancel"
								Else
									BeasongState="3"
									BeasongStateName="배송대기"
									BeasongStateClass="standby"
								End If

								If OrderItemDic.item(iy)(10)="Y" Then
									BeasongState="4"
									BeasongStateName="주문취소"
									BeasongStateClass="odrCancel"
								End If
							End If
							%>
							<% If SameOrderCheck=OrderItemDic.item(ix)(1) And LoopCount=0 And SameOrderCheck2="Y" Then %>
							<li id="<%=OrderItemDic.item(ix)(1)%>_stateclass" class="" onClick="fnAPPpopupOrderBasicInfo('<%=g_AdminURL%>/apps/academy/ordermaster/orderBasicInfo.asp?orderserial=<%= OrderItemDic.item(ix)(1) %>&odiv=<%=odiv%>','<%= OrderItemDic.item(ix)(1) %>')">
								<div class="artStatus">
									<p><span><%= FormatDate(OrderItemDic.item(ix)(0),"0000.00.00") %></span><span>ㅣ</span><span><%= OrderItemDic.item(ix)(1) %></span></p>
									<p class="rt"><span class="nowStatus" id="<%=OrderItemDic.item(ix)(1)%>_state"></span></p>
								</div>
								<div class="artInfo">
									<div class="artThumb"><img src="<%= OrderItemDic.item(ix)(3) %>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
									<strong><% =OrderItemDic.item(ix)(4) %></strong>
									<div class="artTxt">
										<p id="<%=OrderItemDic.item(ix)(1)%>"></p>
									</div>
								</div>
							</li>
							<%
							If BeasongState<>"0" Or BeasongState<>"4" Then
								BadgeCount=BadgeCount+1
							End If
							%>
							<% LoopCount=LoopCount+1 %>
							<% End If %>
							<% Next %>
							<script>
								var itemcount="<%=ItemCount-1%>";
								if(itemcount>0){
									$("#<%=SameOrderSerial%>").empty().append("외 " + itemcount + "건");
								}
								<% If odiv="D" then %>
								$("#<%=SameOrderSerial%>_state").empty().append("<strong><%=BeasongStateName%></strong>");
								$("#<%=SameOrderSerial%>_stateclass").addClass("<%=BeasongStateClass%>");
								<% else %>
									<% if BeasongState=4 then %>
										$("#<%=SameOrderSerial%>_state").empty().append("<strong><%=BeasongStateName%></strong>");
										$("#<%=SameOrderSerial%>_stateclass").addClass("<%=BeasongStateClass%>");
									<% else %>
										$("#<%=SameOrderSerial%>_state").empty().append("<strong>확인대기</strong>");
										$("#<%=SameOrderSerial%>_stateclass").addClass("chkWait");
									<% end if %>
								<% end if %>
							</script>
							<% ItemCount=0 %>
							<% LoopCount=0 %>
							<% 
							SameOrderCheck2="N"
							SameOrderSerial=""
							BeasongCnt=0
							MibeasongCnt=0
							MiChulGoCheck=0
							BeasongState=""
							BeasongStateName=""
							BeasongStateClass=""
							OrderCanCelCnt=0
							%>
							<% Next %>
						</ul>
						<% if ojumun.FTotalCount>ojumun.FPageSize then %>
						<div class="paging">
							<%=fnDisplayPaging_New(page,ojumun.FTotalCount,ojumun.FPageSize,"jsGoPage")%>
						</div>
						<% end if %>
					</div>
				</form>
				</div>
				<% Else %>
				<div class="artNo">
					<div class="linkNotice">
						<% If odiv="S" Then %><p class="fs1-5r">접수된 주문이 없습니다.</p><% End If %>
						<% If odiv="D" Then %><p class="fs1-5r">진행중인 주문이 없습니다.</p><% End If %>
					</div>
				</div>
				<% End If %>
<% Else %>
					<% if ojumun.FresultCount > 0 then %>
					<div class="csListWrap">
						<ul class="artList">
							<% for ix=0 to ojumun.FresultCount-1 %>
							<li class="<% If ojumun.FItemList(ix).Fcurrstate>="B006" Then %>complete<% Else %>undoit<% End If %>" onClick="fnAPPpopupOrderBasicInfo('<%=g_AdminURL%>/apps/academy/ordermaster/csInfo.asp?orderserial=<%= ojumun.FItemList(ix).Forderserial %>&odiv=<%=odiv%>&id=<%=ojumun.FItemList(ix).Fid%>','<%= ojumun.FItemList(ix).Forderserial %>')">
								<div class="artStatus">
									<p><span><%= FormatDate(ojumun.FItemList(ix).Fregdate,"0000.00.00") %></span><span>ㅣ</span><span><%= ojumun.FItemList(ix).Forderserial %></span></p>
									<p class="rt"><span class="nowStatus"><strong><% = ojumun.FItemList(ix).CsStateName %></strong></span></p>
								</div>
								<div class="artInfo">
									<p><span class="tag3"><%= (ojumun.FItemList(ix).Fgubun01Name) %> <i class="arwRt"></i> <%= (ojumun.FItemList(ix).Fgubun02Name) %></span></p>
									<strong><%= ojumun.FItemList(ix).FTitle %></strong>
								</div>
							</li>
							<% Next %>
						</ul>
						<% if ojumun.FTotalCount>ojumun.FPageSize then %>
						<div class="paging">
							<%=fnDisplayPaging_New(page,ojumun.FTotalCount,ojumun.FPageSize,"jsGoPage")%>
						</div>
						<% end if %>
					</div>
				</form>
				</div>
				<% Else %>
				<div class="artNo" style="display:">
					<div class="linkNotice">
						<p class="fs1-5r">수신된 CS 내역이 없습니다.</p>
					</div>
				</div>
				<% End If %>
<% End If %>
			</div>
		</div>
		<!--// content -->

		<!-- 알림 메세지 -->
		<div class="attentionBar" style="display:none" id="alert1">
			<p>출고완료 기록은 6개월간 보관됩니다.</p>
		</div>

		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
<% If statediv=5 and page=1 Then %>
$('#alert1').fadeIn(800).css("display","");
setTimeout(function(){
		$("#alert1").fadeOut(1000);
	}, 5000);
$('#alert1').fadeIn(800).css("display","none");
<% End If %>
setTimeout(function(){
	fnAPPChangeBadgeCount("ordercount",<%=StandByConfirmCnt+BMiBeasongCnt+OrderCSCnt%>);
}, 1000);

});
//-->
</script>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/apps/academy/lib/pms_badge_check.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->