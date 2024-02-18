<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 주문 정보"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
dim MakerID, OrderSerial, ojumun, oitem, ix, odiv, CheckOrderConfirmYN

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
OrderSerial = RequestCheckVar(request("orderserial"),12)
odiv = RequestCheckVar(request("odiv"),1)

set ojumun = New CJumunMaster
ojumun.FRectOrderSerial = OrderSerial
If OrderSerial <> "" Then
ojumun.OneOrderMasterInfo
End If

If MakerID="" Then
	Response.Write "<script>alert('계정 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
ElseIf (OrderSerial="") Then
	Response.Write "<script>alert('주문 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
ElseIf (ojumun.FOneItem.Fidx="") Then
	Response.Write "<script>alert('주문 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

If odiv="S" And ojumun.FOneItem.Fcancelyn<>"Y" Then'주문확인 처리
	if SetUpCheOrderConfirm(OrderSerial, MakerID) > 0 Then
		CheckOrderConfirmYN="Y"
	Else
		CheckOrderConfirmYN="N"
	End If
End If

set oitem = new CJumunMaster
oitem.FRectOrderSerial = OrderSerial
oitem.FRectDesignerID=MakerID
oitem.OrderDetailInfo

'일부 출고 확인
Dim BeasongCnt, BeasongState, MibeasongCnt, MiChulGoCheck
Dim MiChulGoDayCalculate, TotalOrderPrice, OrderCanCelCnt
Dim timelinereturn1, timelinereturn2, CheckOrderCancelView, masteridx
BeasongCnt=0
MibeasongCnt=0
TotalOrderPrice=0
OrderCanCelCnt=0

For ix=0 To oitem.FResultCount-1
	TotalOrderPrice=TotalOrderPrice+oitem.FMasterItemList(ix).Fitemcost
	If oitem.FMasterItemList(ix).Fsongjangno <> "" Then
		BeasongCnt=BeasongCnt+1
	End If
	If oitem.FMasterItemList(ix).Fcode <> "" Then
		MibeasongCnt=MibeasongCnt+1
	End If
	If oitem.FMasterItemList(ix).FCancelYn="Y" Or ojumun.FOneItem.FCancelYn="Y" Then
		OrderCanCelCnt=OrderCanCelCnt+1
	End If
	If (DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now())>0) Then
		MiChulGoCheck=MiChulGoCheck+1
			If ix=0 Then MiChulGoDayCalculate=DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now())
			If MiChulGoDayCalculate > DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now()) Then'최고 미 출고 일 수 가져오기
				MiChulGoDayCalculate=DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now())
			End If
	End If
Next

If oitem.FResultCount>0 Then
	timelinereturn1 = replace(oitem.FMasterItemList(0).FItemname,"&","ø")
	timelinereturn2 = oitem.FResultCount-1
End If

If oitem.FResultCount = BeasongCnt Then
	BeasongState="0"'출고완료
ElseIf oitem.FResultCount > BeasongCnt And BeasongCnt>0 Then
	BeasongState="1"'일부출고
ElseIf MibeasongCnt>0 Or MiChulGoCheck>0 Then
	BeasongState="2"'미출고
ElseIf oitem.FResultCount = OrderCanCelCnt Then
	BeasongState="4"'주문취소
Else
	BeasongState="3"'배송대기
End If

If ojumun.FOneItem.FCancelYn="Y" Then
	BeasongState="4"'주문취소
End If

masteridx = ojumun.FOneItem.Fidx
%>
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// textarea auto size
	$(".searchInput input").keyup(function () {
		$(this).parent().find('button').fadeIn();
	});
});

function fnGoTapPage(param){
	if(param=="B"){
		location.href="/apps/academy/ordermaster/orderBasicInfo.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>";
	}else{
		location.href="/apps/academy/ordermaster/orderTimeline.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>&tl1=<%=timelinereturn1%>&tl2=<%=timelinereturn2%>";
	}
}

function fnCallPhone(phonenum){
	if(confirm("받는분께 연락처로 연결하시겠습니까?")){
		fnAPPpopupOuterBrowser("<%=g_AdminURL%>/apps/academy/ordermaster/popup/callphone.asp?phonenum="+phonenum);
	}
}

function fnSongjangInputCheck(){
	var checkedcnt=0;
	var arrdetailidx='';
	$("input[name=ordercheck]").each(function(i){
		if($("input[name=ordercheck]:eq(" + i + ")").is(":checked")==true){
			if(checkedcnt<1){
				arrdetailidx += $("input[name=ordercheck]:eq(" + i + ")").val();
			}else{
				arrdetailidx += "," + $("input[name=ordercheck]:eq(" + i + ")").val();
			}
			checkedcnt++;
		}
	});
	
	if(checkedcnt<1){
		if($("input[name=ordercheck]").length<1){
			alert("송장번호를 누르면 배송정보를 변경 할 수 있습니다.");
		}else{
			fnSongjangAllCheck();
		}
	}else{
		fnAPPpopupSongjangInput("<%=g_AdminURL%>/apps/academy/ordermaster/popup/popInvoiceWrite.asp?orderserial=<%=OrderSerial%>&arrdetailidx="+arrdetailidx);
	}
}

function fnSongjangAllCheck(){
	var arrdetailidx='';
	$("input").prop('checked', true);
	$("input[name=ordercheck]").each(function(i){
		if(i<1){
			arrdetailidx += $("input[name=ordercheck]:eq(" + i + ")").val();
		}else{
			arrdetailidx += "," + $("input[name=ordercheck]:eq(" + i + ")").val();
		}
	});
	fnAPPpopupSongjangInput("<%=g_AdminURL%>/apps/academy/ordermaster/popup/popInvoiceWrite.asp?orderserial=<%=OrderSerial%>&arrdetailidx="+arrdetailidx);
}

function fnUnDeliverReasonCheck(){
	var checkedcnt=0;
	var arrdetailidx='';
	$("input[name=ordercheck]").each(function(i){
		if($("input[name=ordercheck]:eq(" + i + ")").is(":checked")==true){
			checkedcnt++;
			if(i<1){
				arrdetailidx += $("input[name=ordercheck]:eq(" + i + ")").val();
			}else{
				arrdetailidx += "," + $("input[name=ordercheck]:eq(" + i + ")").val();
			}
		}
	});
	
	if(checkedcnt<1){
		if($("input[name=ordercheck]").length<1){
			alert("출고 완료된 주문은 미출고 사유를 입력 할 수 없습니다.");
		}else{
			fnUnDeliverReasonAllCheck();
		}
	}else{
		fnAPPpopupUnDeliverReasonInput("<%=g_AdminURL%>/apps/academy/ordermaster/popup/popUndeliverReason.asp?orderserial=<%=OrderSerial%>&arrdetailidx="+arrdetailidx);
	}
}

function fnUnDeliverReasonAllCheck(){
	var arrdetailidx='';
	$("input").prop('checked', true);
	$("input[name=ordercheck]").each(function(i){
		if(i<1){
			arrdetailidx += $("input[name=ordercheck]:eq(" + i + ")").val();
		}else{
			arrdetailidx += "," + $("input[name=ordercheck]:eq(" + i + ")").val();
		}
	});
	fnAPPpopupUnDeliverReasonInput("<%=g_AdminURL%>/apps/academy/ordermaster/popup/popUndeliverReason.asp?orderserial=<%=OrderSerial%>&arrdetailidx="+arrdetailidx);
}

function fnOrderListReload(){
	fnAPPParentsWinJsCall("fnThisPageRelold(\"\")");
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden"><%=OrderSerial%></h1>
			<div class="orderManage">
				<ul class="listTab">
					<li class="current"><div>기본정보</div></li>
					<li onclick="fnGoTapPage('T')"><div>타임라인</div></li>
				</ul>

				<div class="artListCont<% If BeasongState=0 Then %> releaseFin<% ElseIf BeasongState=1 Then %> releaseIng<% ElseIf BeasongState=2 Then %> undeliver<% ElseIf ojumun.FOneItem.Fcancelyn="Y" Then %> odrCancel<% Else %> standby<% End If %>">
					<div class="stateWrap">
						<% If BeasongState=0 Then %>
						<div class="stateView">
							<strong>출고완료</strong>
							<p>모든 작품이<br />정상적으로 출고완료 되었습니다.</p>
						</div>
						<% ElseIf BeasongState=1 Then %>
						<div class="stateView">
							<strong>일부출고</strong>
							<p>전체 주문내역 중 일부 작품이 출고 되었습니다.<br />나머지 주문 건도 출고 바랍니다.</p>
						</div>
						<% ElseIf BeasongState=2 Then %>
						<div class="stateView">
							<strong>미출고</strong>
							<p>최초 배송 예정일로부터 <%=MiChulGoDayCalculate%>일이 지났습니다.<br />주문 받은 작품을 배송해주세요.</p>
						</div>
						<% ElseIf BeasongState=4 Or ojumun.FOneItem.Fcancelyn="Y" Then %>
						<div class="stateView">
							<strong>주문취소</strong>
							<p>주문이 취소되었습니다.</p>
						</div>
						<% Else %>
						<div class="stateView">
							<strong>배송대기</strong>
							<p>결제가 완료되었습니다.<br />주문 받은 작품을 배송해주세요.</p>
						</div>
						<% End If %>
						<div class="stateDate">
							<p>
								<span>주문일 : <%=FormatDate(ojumun.FOneItem.Fregdate,"0000.00.00 00:00:00")%></span>
								<span>결제일 : <%=FormatDate(ojumun.FOneItem.Fipkumdate,"0000.00.00 00:00:00")%></span>
								<% If BeasongState=0 Or BeasongState=1 Then %>
								<span>출고일 : <%=FormatDate(ojumun.FOneItem.Fbeadaldate,"0000.00.00 00:00:00")%></span>
								<% End If %>
								<% If BeasongState=4 Or ojumun.FOneItem.Fcancelyn="Y" Then %>
								<span>취소일 : <%=FormatDate(ojumun.FOneItem.Fcanceldate,"0000.00.00 00:00:00")%></span>
								<% End If %>
							</p>
						</div>
						<div class="stateTotal">
							<p>총 합계 : <%=FormatNumber(TotalOrderPrice,0)%>원</p>
						</div>
					</div>
					<div class="orderInfoWrap">
						<h2>주문 작품</h2>
						<form method="post" name="orderfrm">
						<ul class="artList">
							<% If oitem.FResultCount>0 Then %>
							<% For ix=0 To oitem.FResultCount-1 %>
							<% If oitem.FMasterItemList(ix).Fsongjangno<>"" Then %>
							<li class="releaseFin">
							<% ElseIf oitem.FMasterItemList(ix).FCurrstate<7 And (DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now())>0 Or oitem.FMasterItemList(ix).Fcode<>"") Then %>
							<li class="undeliver">
							<% ElseIf oitem.FMasterItemList(ix).FCancelYn="Y" Or (ojumun.FOneItem.Fcancelyn="Y") Then %>
							<li class="odrCancel">
							<% Else %>
							<li class="standby">
							<% End If %>
								<div class="artStatus">
									<% If oitem.FMasterItemList(ix).Fsongjangno<>"" Then %>
									<p><span><%=FormatDate(oitem.FMasterItemList(ix).Fbeasongdate,"0000.00.00")%></span><span>ㅣ</span><span><%=oitem.FMasterItemList(ix).FItemid%></span></p>
									<p class="rt"><span class="invoiceNo" onclick="fnAPPpopupSongjangInput('<%=g_AdminURL%>/apps/academy/ordermaster/popup/popInvoiceWrite.asp?orderserial=<%=OrderSerial%>&arrdetailidx=<%=oitem.FMasterItemList(ix).Fdetailidx%>&mode=edit');"><%=oitem.FMasterItemList(ix).GetSongJangDivName%>&nbsp;<%=oitem.FMasterItemList(ix).Fsongjangno%></span></p>
									
									<% ElseIf oitem.FMasterItemList(ix).FCurrstate<7 And (DateDiff("d",DateAdd("d",oitem.FMasterItemList(ix).Frequiremakeday+2,oitem.FMasterItemList(ix).Fupcheconfirmdate),now())>0 Or oitem.FMasterItemList(ix).Fcode<>"") Then %>
									<p class="lPad0-5r" style="width:3.4rem;"><input type="checkbox" id="ordercheck" name="ordercheck" value="<%=oitem.FMasterItemList(ix).Fdetailidx%>" /></p>
									<p><%=oitem.FMasterItemList(ix).FItemid%></p>
									<p class="rt"><span class="nowStatus"><% If oitem.FMasterItemList(ix).Fcode<>"" Then %><%=oitem.FMasterItemList(ix).getMiSendCodeName%><% Else %>미출고<% End If %></span></p>
									
									<% ElseIf oitem.FMasterItemList(ix).FCancelYn="Y" Or (ojumun.FOneItem.Fcancelyn="Y") Then %>
									<p><span><% If oitem.FMasterItemList(ix).Fcanceldate<>"" Then %><%=FormatDate(oitem.FMasterItemList(ix).Fcanceldate,"0000.00.00")%><% Else %><%=FormatDate(ojumun.FOneItem.Fcanceldate,"0000.00.00")%><% End If %></span><span>ㅣ</span><span><%=oitem.FMasterItemList(ix).FItemid%></span></p>
									<p class="rt"><span class="nowStatus"><strong>주문취소</strong></span></p>
									
									<% Else %>
									<p class="lPad0-5r" style="width:3.4rem;"><input type="checkbox" id="ordercheck" name="ordercheck" value="<%=oitem.FMasterItemList(ix).Fdetailidx%>" /></p>
									<p><%=oitem.FMasterItemList(ix).FItemid%></p>
									<p class="rt"><span class="nowStatus"><strong>배송대기</strong></span></p>
									<% End If %>
								</div>
								<div class="artInfo">
									<div class="artThumb"><img src="<%=oitem.FMasterItemList(ix).FListimage%>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
									<strong><%=oitem.FMasterItemList(ix).FItemname%></strong>
									<div class="artTxt">
										<p><dfn><%=oitem.FMasterItemList(ix).Fitemoptionname%></dfn></p>
										<p><dfn><%=oitem.FMasterItemList(ix).Fitemno%>개</dfn></p>
										<p class="tPad1r"><span class="salePrice"><%=FormatNumber((oitem.FMasterItemList(ix).Fitemcost*oitem.FMasterItemList(ix).Fitemno),0)%>원</span></p>
									</div>
								</div>
								<% If oitem.FMasterItemList(ix).Frequiredetail<>"" Then %>
								<div class="boxUnit bdrTRtGry">
									<div class="boxHead">
										<b>주문제작 메시지</b>
									</div>
									<div class="boxCont"><%=oitem.FMasterItemList(ix).Frequiredetail%></div>
								</div>
								<% End If %>
								<% If oitem.FMasterItemList(ix).Fipgodate<>"" Then %>
								<div class="boxUnit shortCont"><strong class="fs1-2r">배송 예정일 : <%=FormatDate(oitem.FMasterItemList(ix).Fipgodate,"0000년00월00일")%></strong></div>
								<% End If %>
							</li>
							<% Next %>
							<% End If %>
						</ul>
						</form>
						<h2>구매/배송 정보</h2>
						<ul class="artList dfCompos2">
							<li>
								<ul class="list">
									<li>
										<dfn class="cGy1"><b>주문자</b></dfn>
										<div class="cGy1"><%=ojumun.FOneItem.Fbuyname%></div>
									</li>
									<li>
										<dfn class="cGy1"><b>주문자 연락처</b></dfn>
										<div class="cGy1">
											<% If ojumun.FOneItem.Fbuyphone <> "" And ojumun.FOneItem.Fbuyphone <> "--" Then %><p><%=ojumun.FOneItem.Fbuyphone%></p><% End If %>
											<% If ojumun.FOneItem.Fbuyhp <> ""  And ojumun.FOneItem.Fbuyhp <> "--" Then %><p><%=ojumun.FOneItem.Fbuyhp%></p><% End If %>
										</div>
									</li>
								</ul>
							</li>
							<li>
								<ul class="list">
									<li>
										<dfn><b>받는분</b></dfn>
										<div class="cGy3"><%=ojumun.FOneItem.Freqname%></div>
									</li>
									<li>
										<dfn><b>받는분 연락처</b></dfn>
										<div class="cGy3">
											<% If ojumun.FOneItem.Freqphone <> "" And ojumun.FOneItem.Freqphone <> "--" Then %><p><a href="javascript:fnCallPhone('<%=ojumun.FOneItem.Freqphone%>');" class="txtUdrLine"><%=ojumun.FOneItem.Freqphone%></a></p><% End If %>
											<% If ojumun.FOneItem.Freqhp <> ""  And ojumun.FOneItem.Freqhp <> "--" Then %><p><a href="javascript:fnCallPhone('<%=ojumun.FOneItem.Freqhp%>');" class="txtUdrLine"><%=ojumun.FOneItem.Freqhp%></a></p><% End If %>
										</div>
									</li>
									<li>
										<dfn><b>받는분 주소</b></dfn>
										<div class="cGy3">[<%=ojumun.FOneItem.Freqzipcode%>] <%=ojumun.FOneItem.Freqzipaddr%> <%=ojumun.FOneItem.Freqaddress%></div>
									</li>
									<% If ojumun.FOneItem.Fbeasongmemo<>"" Then %>
									<li>
										<dfn><b>기타사항</b></dfn>
										<div class="cGy3"><%=ojumun.FOneItem.Fbeasongmemo%></div>
									</li>
									<% End If %>
								</ul>
							</li>
						</ul>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->

		<!-- 알림 메세지 -->
		<div class="attentionBar" style="display:none" id="alert1">
			<p>주문 확인 되었습니다. 해당 주문내역은 주문처리 탭에서 확인 가능합니다.</p>
		</div>
		<div class="attentionBar" style="display:none" id="alert2">
			<p>주문취소가 확인 되었습니다. 해당 취소내역은 '관련 CS' 탭에서 확인 가능합니다.</p>
		</div>
		<% If BeasongState=4 Or ojumun.FOneItem.Fcancelyn="Y" Or BeasongState=0 Then %>
		<% Else %>
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnWishV16a" onClick="fnSongjangInputCheck();">출고처리</button></p>
			<p><button type="button" class="btnV16a btnRed2V16a" onClick="fnUnDeliverReasonCheck();">미출고 사유 입력</button></p>
		</div>
		<!-- 하단 플로팅 버튼 -->
		<div id="layerMask" class="layerMask"></div>
		<% End If %>
	</div>
</div>
</body>
</html>
<%
If BeasongState="4" Then
	CheckOrderCancelView = GetOrderCancelViewCheck(masteridx)
End If
%>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
<% If CheckOrderConfirmYN="Y" Then %>
fnAPPParentsWinReLoad();
$('#alert1').fadeIn(800).css("display","");
setTimeout(function(){
		$("#alert1").fadeOut(1000);
	}, 5000);
$('#alert1').fadeIn(800).css("display","none");
<% End If %>
<% If CheckOrderCancelView="Y" Then %>
$('#alert2').fadeIn(800).css("display","");
setTimeout(function(){
		$("#alert2").fadeOut(1000);
	}, 5000);
$('#alert2').fadeIn(800).css("display","none");
<% End If %>
});
//-->
</script>
<%
Set ojumun = Nothing
Set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->