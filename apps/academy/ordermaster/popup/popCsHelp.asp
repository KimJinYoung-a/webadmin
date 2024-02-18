<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - CS 관련 도움말"
%>
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
dim divcd
divcd = RequestCheckVar(request("divcd"),5)
%>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
		<div class="content">
			<h1 class="hidden">CS 관련 도움말</h1>
			<!-- for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
			<div class="csHelp">
				<div class="boxUnit bdrTRtGry">
					<% if divcd="A000" then %> <!-- 맞교환 설명 -->
						<b>* 맞교환 도움말</b>
					<% elseif divcd="A001" then %> <!-- 누락재발송 설명 -->
						<b>* 누락재발송 도움말</b>
					<% elseif divcd="A004" then %> <!-- 반품 설명 -->
						<b>* 반품관련 도움말</b>
						<br>반품접수가 될경우, 고객님께 발송하신 택배사 전화번호를 안내해드리며,
						<br>상품을 받으신 택배사를 통해 착불반송을 해주시도록 안내를 해드리고 있습니다.
						<br>변심 반품의 경우, 착불반송포함 왕복배송비를 차감한 금액을 고객님께 환불해드리며,
						<br>차감된 금액은 업체정산내역에 자동으로 등록됩니다.
						<br>(편도 2,000원 / 왕복 4,000원 차감)
						<br>
						<br>반송상품이 도착하면, 접수내용을 확인하신 후,
						<br>아래쪽 처리내용에 내용을 남겨주시면, 고객센터에 내용이 전달되며,
						<br>고객센터에서 반품취소처리 및 고객환불을 진행합니다.
						<br>
						<br>*처리프로세스
						<br>1.접수
						<br>2.업체완료처리 --> 고객센터에 처리결과 전달
						<br>3.고객센터완료처리 --> 고객에게 처리결과 안내 및 메일발송
					<% elseif divcd="A006" then %> <!-- 출고시 유의사항 설명 -->
						<b>* 출고시 유의사항 도움말</b>
						<br>주문건 확인 후, 고객님이 주문관련 변경을 요청하셨을 경우,
						<br>출고시 유의사항으로 등록됩니다.
						<br>ex)배송지변경/상품변경/상품옵션변경
						<br>
						<br>텐바이텐 고객센터에서 별도로 가능여부 확인을 위해 연락드립니다.
					<% else %>

					<% end if %>
				</div>
			</div>
			<!-- //for dev msg : 레이어 작업시 이 부분만 가져가면 됩니다.-->
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>