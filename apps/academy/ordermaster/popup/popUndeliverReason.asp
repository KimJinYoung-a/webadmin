<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 송장입력"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/ordermaster/ordercls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
dim MakerID, OrderSerial, oitem, ix, ordercheck, mode

MakerID = requestCheckVar(request.cookies("partner")("userid"),32)
OrderSerial = RequestCheckVar(request("orderserial"),12)
mode = RequestCheckVar(request("mode"),12)
ordercheck = requestCheckVar(request("arrdetailidx"),128)

'Response.write ordercheck
'Response.end

set oitem = new CJumunMaster
oitem.FRectDetailIDx = ordercheck
oitem.OrderDetailInfoInidx

If oitem.FMasterItemList(0).Fcode<>"" Then mode="edit"
%>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css" />
<script>
$(function() {
	// button tab
	$(".selectBtn button").click(function(){
		$(this).parent().parent().find("button").removeClass("selected");
		$(this).addClass("selected");
	});

	// datepicker
	$("#ipgodate").datepicker({
		showOn:"both",
		buttonImage: "http://image.thefingers.co.kr/apps/2016/ico_calendar.png",
		buttonImageOnly:true,
		buttonText:"출고 예정일을 선택해주세요",
		dateFormat:"yy-mm-dd"
	});
});

function fnMisendReason(reason){
	if(reason!="05"){
		$("#IpgoDate").css("display","");
	}else{
		$("#IpgoDate").css("display","none");
	}
	$("#MisendReason").val(reason);
}

function fnitemSoldOutFlag(flag){
	$("#itemSoldOut").val(flag);
}

function fnAppCallWinConfirm(){
    var frm = document.frmMisend;
    var today= new Date();
    //today = new Date(today.getYear(),today.getMonth(),today.getDate());  //오늘도 가능하도록
    today = new Date(<%=year(now())%>,<%=month(now())-1%>,<%=Day(now())%>);  //2016/09/08 수정.
    
    var inputdate;

	if(frm.mode.value=="edit"){
		alert('미출고 사유는 수정 할 수 없습니다.');
        return;
	}
    if (frm.MisendReason.value.length<1){
        alert('미출고 사유를 선택해주세요.');
        frm.MisendReason.focus();
        return;
    }
    //출고지연(03), 주문제작(02), 예약배송(04)
    if ((frm.MisendReason.value=="03")||(frm.MisendReason.value=="02")||(frm.MisendReason.value=="04")){
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('출고 예정일을 입력하세요.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('출고 예정일은 오늘 이후날짜로 설정이 가능합니다.');
            //ipgodate.focus();
            return;
        }
    }
	if(frm.MisendReason.value!="05"){
		if (confirm('선택하신 미출고 사유의 내용으로 고객님께\nSMS와 이메일이 발송됩니다.')){
			frm.mode.value="misendInput";
			frm.target="FrameCKP";
			frm.action = "domisendinput.asp";
			frm.submit();
		}
	}else{
		if (confirm('품절로 인한 출고불가일 경우, 더핑거스\n고객센터에서 별도로 고객님께 연락을 드립니다.')){
			frm.mode.value="misendInput";
			frm.target="FrameCKP";
			frm.action = "domisendinput.asp";
			frm.submit();
		}
	}
}

function fnMisendReasonInputEnd(){
	alert("입력이 완료되었습니다.");
	fnAPPParentsWinReLoad();
	setTimeout(function(){
		fnAPPclosePopup();
	}, 300);
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="frmMisend">
		<input type="hidden" name="orderserial" value="<%=OrderSerial%>">
		<input type="hidden" name="mode" value="<%=mode%>">
		<input type="hidden" name="MisendReason" id="MisendReason">
		<input type="hidden" name="itemSoldOut" id="itemSoldOut">
		<div class="content bgGry">
			<h1 class="hidden">미출고 사유 입력</h1>
			<div class="unDeliverReason">
				<ul class="artList">
					<li class="tPad0-5r">
						<div class="boxUnit bdrTRtGry">
							<div class="boxCont cGy2 fs1-2r" style="text-indent:-1.3rem; padding-left:1.3rem">
								<p>1. 미출고 사유가 출고지연 및 주문제작일 경우, 아래의 내용으로 고객님께 SMS와 이메일이 발송됩니다.</p>
								<p class="tPad0-5r">2. 고객님께 안내된 출고예정일을 꼭 지켜주시기 바라며, 변동사항이 생길 경우, 고객센터로 연락 부탁드립니다.</p>
								<p class="tPad0-5r">3. 품절출고불가인 경우, 더핑거스 고객센터에서 별도로 고객님께 연락을 드릴 예정입니다.</p>
							</div>
						</div>
					</li>
					<% If oitem.FResultCount>0 Then %>
					<% For ix=0 To oitem.FResultCount-1 %>
					<li>
						<div class="artInfo">
							<input type="hidden" name="detailidx" value="<%=oitem.FMasterItemList(ix).Fdetailidx%>">
							<input type="hidden" name="Sitemid" value="<%= oitem.FMasterItemList(ix).FItemID %>">
							<input type="hidden" name="Sitemoption" value="<%= oitem.FMasterItemList(ix).FItemOption %>">
							<div class="artThumb"><img src="<%=oitem.FMasterItemList(ix).FListimage%>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
							<p class="orderNo"><%=oitem.FMasterItemList(ix).FItemid%></p>
							<strong><%=oitem.FMasterItemList(ix).FItemname%></strong>
							<div class="artTxt">
								<p><dfn><%=oitem.FMasterItemList(ix).Fitemoptionname%></dfn></p>
								<p><dfn><%=oitem.FMasterItemList(ix).Fitemno%>개</dfn></p>
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
					</li>
					<% Next %>
					<% End If %>
				</ul>
				<div class="registUnit">
					<h2><b>미출고 사유 선택</b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid4"><button type="button" class="btnM1 btnGry<% If oitem.FMasterItemList(0).Fcode="03" Then %> selected<% End If %>" onClick="fnMisendReason('03');">출고지연</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry<% If oitem.FMasterItemList(0).Fcode="05" Then %> selected<% End If %>" onClick="fnMisendReason('05');">품절</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry<% If oitem.FMasterItemList(0).Fcode="02" Then %> selected<% End If %>" onClick="fnMisendReason('02');">주문제작</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry<% If oitem.FMasterItemList(0).Fcode="04" Then %> selected<% End If %>" onClick="fnMisendReason('04');">예약배송</button></div>
						</li>
					</ul>
					<ul class="list">
						<li class="selectBtn">
							<div class="grid4"><button type="button" class="btnM1 btnGry selected" onClick="fnitemSoldOutFlag('N');">상품 품절처리</button></div>
							<div class="grid4"><button type="button" class="btnM1 btnGry" onClick="fnitemSoldOutFlag('S');">상품 일시품절처리</button></div>
						</li>
					</ul>
				</div>
				<div class="registUnit" id="IpgoDate" style="display:<% If oitem.FMasterItemList(0).Fcode<>"05" And oitem.FMasterItemList(0).Fcode<>"" Then %><% Else %>none<% End If %>">
					<h2 class=""><b><a href="http://code.jquery.com/ui/1.12.1/themes/smoothness/jquery-ui.css">출고 예정일</a></b></h2>
					<ul class="list">
						<li class="selectBtn">
							<div class="datePickWrap"><input type="text" name="ipgodate" value="<%=oitem.FMasterItemList(0).Fipgodate%>" id="ipgodate" readonly placeholder="출고 예정일을 선택해주세요"></div>
						</li>
					</ul>
				</div>
				
			</div>
		</div>
		</form>
		<!--// content -->
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
fnAPPShowRightConfirmBtns();
});
//-->
</script>
<%
Set oitem = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->