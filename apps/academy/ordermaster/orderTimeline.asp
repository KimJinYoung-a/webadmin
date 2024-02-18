<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 타임 라인"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
Dim MakerID, OrderSerial, odiv, ojumun, itemname, itemcnt, ix
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
OrderSerial = RequestCheckVar(request("orderserial"),12)
odiv = RequestCheckVar(request("odiv"),1)
itemname = RequestCheckVar(replace(request("tl1"),"ø","&"),32)
itemcnt = RequestCheckVar(request("tl2"),2)

If (OrderSerial="" Or MakerID="") Then
	Response.Write "<script>alert('주문 정보가 없습니다.');fnAPPclosePopup();</script>"
	Response.End
End If

set ojumun = New CJumunMaster
ojumun.FRectOrderSerial = OrderSerial
ojumun.FRectDesignerID = MakerID
ojumun.OrderTimeLineList

%>
<script>
<!--
function fnGoTapPage(param){
	if(param=="B"){
		location.href="/apps/academy/ordermaster/orderBasicInfo.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>";
	}else{
		location.href="/apps/academy/ordermaster/orderTimeline.asp?orderserial=<%=OrderSerial%>&odiv=<%=odiv%>";
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
					<li onclick="fnGoTapPage('B')"><div>기본정보</div></li>
					<li class="current"><div>타임라인</div></li>
				</ul>

				<div class="timelineWrap">
					<div class="orderInfoBox">
						<span class="odrNo"><b><%=OrderSerial%></b></span>
						<span class="odrName"><dfn><%=itemname%></dfn></span>
						<% If itemcnt > 0 Then %>
						<span class="etc">외 <%=itemcnt%>건</span>
						<% End If %>
					</div>

					<div class="timeline">
						<ul>
							<% If ojumun.FResultCount>0 Then %>
							<% For ix=0 To ojumun.FResultCount-1 %>
							<li class="<%=ojumun.FMasterItemList(ix).StateClassName%><% If ix=0 Then %> current<% End If %>">
								<time><p class="cGy2"><%=FormatDate(ojumun.FMasterItemList(ix).Fregdate,"0000.00.00")%></p><p><%=FormatDate(ojumun.FMasterItemList(ix).Fregdate,"00:00:00")%></p></time>
								<div class="timeCont">
									<p><%=ojumun.FMasterItemList(ix).Fstatediv%></p>
									<% If ojumun.FMasterItemList(ix).Fbeasongetc<>"" Then %>
									<p class="fs1-1r cGy5 tPad0-5r"><%=ojumun.FMasterItemList(ix).Fbeasongetc%></p>
									<% End If %>
								</div>
							</li>
							<% Next %>
							<% End If %>
						</ul>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->