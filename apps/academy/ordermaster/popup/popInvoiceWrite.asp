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


'' 택배사 일괄적용
Sub drawSelectBoxDeliverCompanyAssign(selectBoxName,selectedId)
   dim tmp_str,query1
   %><select class="select" name="<%=selectBoxName%>">
     <option value='' <%if selectedId="" then response.write " selected"%>>택배사를 선택해 주세요</option><%
   query1 = " select top 100 divcd,divname from [db_order].[dbo].tbl_songjang_div where isUsing='Y' "
   query1 = query1 + " order by divcd"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("divcd")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("divcd")&"' "&tmp_str&">" & "" & replace(db2html(rsget("divname")),"'","") &  "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
<script>
<!--
function fnAppCallWinConfirm(){
	var checksongjang=true;
	$(".invoiceWrite input[name=songjangno]").each(function(i){
		if($(".invoiceWrite select[name=songjangdiv]:eq(" + i + ")").val()==""){
			alert("택배사를 선택해 주세요.");
			checksongjang=false;
			return false;
		}else if($(".invoiceWrite input[name=songjangno]:eq(" + i + ")").val()==""){
			alert("운송장 번호를 입력해 주세요.");
			checksongjang=false;
			return false;
		}
	});
	if(checksongjang){
		if(confirm("선택한 주문건을 모두 출고처리하시겠습니까?")){
			document.sfrm.action="/apps/academy/ordermaster/popup/dosongjanginput.asp";
			document.sfrm.mode.value="reg";
			document.sfrm.target="FrameCKP";
			document.sfrm.submit();
		}
	}
}

function fnSongjangInputEnd(OrderStateNum){
	alert("선택된 주문건이 모두 출고 처리되었습니다");
	setTimeout(function(){
		fnAPPParentsWinReLoad();
	}, 300);
	setTimeout(function(){
		fnAPPChangeBadgeCount("ordercount",OrderStateNum);
	}, 600);
	setTimeout(function(){
		fnAPPopenerJsCallClose("fnOrderListReload(\"\")");
	}, 900);
}
//-->
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">송장 입력</h1>
			<div class="invoiceWrite">
				<form method="post" name="sfrm">
				<input type="hidden" name="orderserial" value="<%=OrderSerial%>">
				<input type="hidden" name="mode" value="<%=mode%>">
				<% If oitem.FResultCount>0 Then %>
				<% For ix=0 To oitem.FResultCount-1 %>
				<ul class="artList">
					<li>
						<div class="artInfo">
							<div class="artThumb"><img src="<%=oitem.FMasterItemList(ix).FListimage%>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
							<p class="orderNo"><%=oitem.FMasterItemList(ix).FItemid%></p>
							<strong><%=oitem.FMasterItemList(ix).FItemname%><input type="hidden" name="detailidx" value="<%=oitem.FMasterItemList(ix).Fdetailidx%>"></strong>
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
					<li class="selectBtn">
						<% drawSelectBoxDeliverCompanyAssign "songjangdiv", oitem.FMasterItemList(ix).Fsongjangdiv %>
					</li>
					<li class="list">
						<dfn><b>운송장 번호</b></dfn>
						<div><input type="number" name="songjangno" id="songjangno" value="<%=oitem.FMasterItemList(ix).Fsongjangno%>" placeholder="운송장 번호를 입력해주세요" pattern="[0-9]*" inputmode="numeric" min="0" /></div>
					</li>
				</ul>
				<% Next %>
				<% End If %>
				</form>
			</div>
		</div>
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