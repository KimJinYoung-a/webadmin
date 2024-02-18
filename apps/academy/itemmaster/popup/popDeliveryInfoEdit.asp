<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"

Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 배송비 안내"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Dim itemid
itemid = requestCheckVar(request("itemid"),10)

Dim oitem, ordercomment
set oitem = new CItem
oitem.FRectMakerId = request.cookies("partner")("userid")
oitem.FRectItemID = itemid
if (itemid<>"") then
oitem.GetOneItem
ordercomment=oitem.FOneItem.Fordercomment
End If
%>
<script>
jQuery(document).ready(function(){
	//$("#ordercomment").keyup(function(){
		fnAPPShowRightConfirmBtns();
	//});
});
function fnAppCallWinConfirm(){
	if($("#ordercomment").val()==""){
		alert("배송비 안내를 입력해주세요.");
	}
	else{
		document.sform.action="/apps/academy/itemmaster/popup/DIYItemPopupDetailinfoEdit_Process.asp";
		document.sform.target="FrameCKP";
		document.sform.submit();	
	}
}
function fnDetailInfoEnd(){
    var ordercommenttxt = Base64.encode($("#ordercomment").val());
	//var ordercommenttxt = encSpecialCharNativeFun($("#ordercomment").val());
	fnAPPopenerJsCallClose("fnDeliveryInfoSet(\""+ordercommenttxt+"\")");
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form method="post" name="sform" autocomplete="off">
		<input type="hidden" name="itemid" value="<%=itemid%>">
		<div class="content bgGry">
			<h1 class="hidden">배송비 안내 입력</h1>
			<div class="vodAdd">
				<div class="linkInsert">
					<textarea rows="5" name="ordercomment" id="ordercomment" placeholder="예) 기간, 비용 등"><%=ordercomment%></textarea>
				</div>
				<div class="linkNotice">
					<p class="fs1-5r">구매 고객이 알아야 할 <br />특별한 배송사항을 입력해주세요.</p>
				</div>
			</div>
		</div>
		</form>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<span class="setContView">제품은 배송시 안전을 위해 배송비가 부과됩니다. 제품은 배송시 안전을 위해 배송비가 부과됩니다.</span>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->