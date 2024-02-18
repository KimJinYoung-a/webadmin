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
dim MakerID, OrderSerial, ocs, oitem, ix, odiv, CheckOrderConfirmYN, CSID

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
OrderSerial = RequestCheckVar(request("orderserial"),12)
CSID = RequestCheckVar(request("id"),12)
odiv = RequestCheckVar(request("odiv"),1)
If odiv="" Then odiv="C"

set ocs = New CCSASList
ocs.FRectOrderSerial=OrderSerial
ocs.FRectCSID=CSID
ocs.GetCSASMasterInfo

If (OrderSerial="" Or MakerID="") Then
	Response.Write "<script>alert('CS 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

set oitem = new CJumunMaster
oitem.FRectMasterIDX = CSID
oitem.FRectDesignerID=MakerID
oitem.GetCSASDetailInfo

%>
<script>
<!--
function fnGoTapPage(param){
	if(param=="B"){
		location.href="/apps/academy/ordermaster/csBasicInfo.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>&id=<%=CSID%>";
	}else{
		location.href="/apps/academy/ordermaster/csInfo.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>&id=<%=CSID%>";
	}
}
//-->
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden"><%=OrderSerial%></h1><!-- for dev msg : 주문번호 표기 해주세요 -->
			<div class="orderManage">
				<ul class="listTab">
					<li class="current" onclick="fnGoTapPage('C')"><div>CS정보</div></li>
					<li onclick="fnGoTapPage('B')"><div>주문상세</div></li>
				</ul>

				<div class="artListCont<% If ocs.FOneItem.Fcurrstate>="B006" Then %> finished<% Else %> unDoit<% End If %>"><!-- for dev msg : 처리완료(finished) / 미처리(unDoit) //-->
					<div class="stateWrap">
						<div class="stateView">
							<strong><%=ocs.FOneItem.Fcurrstatename%></strong>
							<p><span class="tag4"><%= (ocs.FOneItem.Fgubun01Name) %> <i class="arwRt"></i> <%= (ocs.FOneItem.Fgubun02Name) %></span></p>
							<p><%= ocs.FOneItem.FTitle %></p>
						</div>
						<div class="stateDate">
							<p>
								<dfn><%= ocs.FOneItem.Fcustomername %>(<%= ocs.FOneItem.Fuserid %>)</dfn>
								<span>작성일 : <%=FormatDate(ocs.FOneItem.Fregdate,"0000.00.00 00:00:00") %></span>
								<% If ocs.FOneItem.Fcurrstate>="B006" Then %>
								<span>완료일 : <%=FormatDate(ocs.FOneItem.Ffinishdate,"0000.00.00 00:00:00") %></span>
								<% End If %>
							</p>
						</div>
					</div>
					<div class="orderInfoWrap">
						<h2>접수 내용</h2>
						<div class="artList"><%= ocs.FOneItem.Fcontents_jupsu %></div>
						<h2>주문 작품</h2>
						<ul class="artList">
							<% If oitem.FResultCount>0 Then %>
							<% For ix=0 To oitem.FResultCount-1 %>
							<li>
								<a href="">
									<div class="artStatus">
										<p><%=oitem.FMasterItemList(ix).FItemid%></p>
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
								</a>
							</li>
							<% Next %>
							<% End If %>
						</ul>
						<h2>CS 처리결과</h2>
						<ul class="artList dfCompos2">
							<li>
								<div class="csResultTxt">
									<%=nl2br(ocs.FOneItem.Fcontents_finish)%>
								</div>
							</li>
							<li>
								<ul class="list">
									<li>
										<dfn><b>관련 운송장</b></dfn>
										<div><%=ocs.FOneItem.Fsongjangdivname%>&nbsp;&nbsp;<%=ocs.FOneItem.Fsongjangno%></div>
									</li>
								</ul>
							</li>
						</ul>
						<div class="addBtn">
							<button type="button" class="btnB1 btnDkGry" onClick="fnAPPpopupCsHelpInfo('<%=g_AdminURL%>/apps/academy/ordermaster/popup/popCsHelp.asp?divcd=<%=ocs.FOneItem.Fdivcd%>');"><span class="question">CS 관련 도움말</span></button>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<!-- 하단 플로팅 버튼 -->
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnWishV16a" onClick="fnAPPpopupCSResultInput('<%=g_AdminURL%>/apps/academy/ordermaster/popup/popCsResult.asp?idx=<%=ocs.FOneItem.Fid%>')">CS 처리결과 작성</button></p>
		</div>

		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<%
set ocs = Nothing
set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->