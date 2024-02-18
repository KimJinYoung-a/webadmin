<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 상품정보제공고시"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Dim itemid
itemid = requestCheckVar(request("itemid"),10)

Dim oitem, infoDiv
set oitem = new CItem
oitem.FRectMakerId = request.cookies("partner")("userid")
oitem.FRectItemID = itemid
if (itemid<>"") then
oitem.GetOneItem
infoDiv=Trim(oitem.FOneItem.FinfoDiv)
End If
%>
<script type="text/javascript" src="/apps/academy/lib/confirm.js"></script>
<script>

function chgInfoDiv(v) {
	$("#infoList").empty();
	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "/apps/academy/itemmaster/act_itemInfoDivFormEdit.asp",
			data: "itemid=<%=itemid%>&ifdv="+v+"&fingerson=on",
			dataType: "html",
			async: false
		}).responseText;

		if(str!="") {
			$("#infoList").html(str);
		}
	}
//	if(v=="35") {
//		$("#lyItemSrc").show();
//		$("#lyItemSize").show();
//	} else {
//		$("#lyItemSrc").hide();
//		$("#lyItemSize").hide();
//	}
}
//단순 라디오 선택자
function chgInfoChk(fm,cnt,yn) {
	$(fm).parent().parent().find("button").removeClass("selected");
	$(fm).addClass("selected");
	$("#info_"+cnt+" input[name=infoChk]").val(yn);
}

//문구 라디오 선택자
function chgInfoSel(fm,cnt,yn) {
	$(fm).parent().parent().find("button").removeClass("selected");
	$("#info_"+cnt+" textarea[name=infoCont]").val($(fm).attr("msg"));
	$(fm).addClass("selected");
	$("#info_"+cnt+" input[name=infoChk]").val(yn);
	if(yn=="Y"){
		$("#info_"+cnt+" textarea[name=infoCont]").removeAttr("readonly");
	}else{
		$("#info_"+cnt+" textarea[name=infoCont]").attr("readonly", true);
	}
}
//문구 라디오 선택자2
function chgInfoSel2(fm,cnt,yn) {
	$(fm).parent().parent().find("button").removeClass("selected");
	$(fm).addClass("selected");
	$("#info_"+cnt+" input[name=infoChk]").val(yn);
	if(yn=="Y"){
		$("#info_"+cnt+" input[name=infoCont]").removeAttr("readonly");
	}else{
		$("#info_"+cnt+" input[name=infoCont]").attr("readonly", true);
	}
}

function fnAppCallWinConfirm(){
	var arrinfochk='', arrinfocont='', arrinfocd='';
	var jsontxt;
	var iteminfo = document.iteminfo;
	if($("#infoDiv option:selected").val()==""){
		alert("품목을 선택해주세요.");
	}else if (validate(iteminfo)==false) {
        return;
    }
	else{
		document.iteminfo.action="/apps/academy/itemmaster/popup/DIYItemPopupDetailinfoEdit_Process.asp";
		document.iteminfo.target="FrameCKP";
		document.iteminfo.submit();
	}
}

function fnDetailInfoEnd(){
	jsontxt = $("#infoDiv").val() + "!" + $("#infoDiv option:selected").text();
	fnAPPopenerJsCallClose("fnItemInfoDivSet(\"" + jsontxt + "\")");
}

jQuery(document).ready(function(){
	<% if infoDiv="" Then %>
	$("#infoDiv").change(function(){
		fnAPPShowRightConfirmBtns();
	});
	<% else %>
		fnAPPShowRightConfirmBtns();
		chgInfoDiv('<%=infoDiv%>');
		$("#infoDiv").val('<%=infoDiv%>').attr("selected","selected");
	<% end if %>
});
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<form name="iteminfo" method="post" onsubmit="return false;" autocomplete="off">
		<input type="hidden" name="itemid" value="<%=itemid%>">
		<input type="hidden" name="mode" value="iteminfo">
		<div class="content bgGry">
			<h1 class="hidden">상품정보 제공 고시</h1>
			<div class="artInfoSet">
				<div class="selectBtn">
					<div>
						<select name="infoDiv" id="infoDiv" onChange="chgInfoDiv(this.value)">
							<option value="">품목을 선택해주세요</option>
							<option value="01">의류</option>
							<option value="02">구두/신발</option>
							<option value="03">가방</option>
							<option value="04">패션잡화(모자/벨트/액세서리)</option>
							<option value="05">침구류/커튼</option>
							<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
							<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
							<option value="17">주방용품</option>
							<option value="18">화장품</option>
							<option value="19">귀금속/보석/시계류</option>
							<option value="20">식품(농수산물)</option>
							<option value="21">가공식품</option>
							<option value="22">건강기능식품/체중조절식품</option>
							<option value="23">영유아용품</option>
							<option value="24">악기</option>
							<option value="25">스포츠용품</option>
							<option value="26">서적</option>
							<option value="35">기타</option>
						</select>
					</div>
				</div>
				<div id="itemInfoCont" style="display:none">
					<ul class="infoList" id="infoList"></ul>
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
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitem = nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->