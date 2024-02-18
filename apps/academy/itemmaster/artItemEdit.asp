<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Session.codepage="65001"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 작품 수정"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
Dim vImgURL, itemid, optionlevel, makerid
vImgURL=""
itemid = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If

Dim npartner
set npartner = new CPartnerUser
npartner.FRectDesignerID = makerid
npartner.GetAcademyPartnerList

dim oitem, i, KeyWordCnt
itemid = RequestCheckVar(request("itemid"),10)

set oitem = new CItem
oitem.FRectMakerId = makerid
oitem.FRectItemID = itemid
if (oitem.FRectItemID<>"") then
oitem.GetOneItem
End If

If oitem.FOneItem.Fkeywords <> "" Then
KeyWordCnt = ubound(Split(oitem.FOneItem.Fkeywords,","))
End If

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
oitemoption.GetItemOptionInfo
end If

If oitemoption.FResultCount > 0 Then
	If oitemoption.IsMultipleOption Then
		optionlevel="2"
	ElseIf oitemoption.FResultCount > 0 And oitemoption.IsMultipleOption=False Then
		optionlevel="1"
	End If
Else
	optionlevel="0"
End If

Dim optlimit
set optlimit = new CItemOption
optlimit.FRectItemID = itemid
if itemid<>"" then
optlimit.GetItemOptionLimitNoInfo
end If

Dim ovod, videoFullUrl
set ovod = new CItem
ovod.FRectMakerId = makerid
ovod.FRectItemID = itemid
if (ovod.FRectItemID<>"") then
ovod.GetItemContentsVideo
videoFullUrl=ovod.FOneItem.FvideoFullUrl
End If

dim oitemimg
set oitemimg = new CItemAddImage
oitemimg.FRectItemID = itemid
if itemid<>"" then
oitemimg.GetOneItemAddImageList
end If
%>
<script>
<!--
function fnVodLinkSet(vodlink){
	vodlink = vodlink.replace(/&dbqt;/g,"\"");
	vodlink = vodlink.replace(/&sgqt;/g,"\'");

	$("#itemvideo").val(vodlink);
}
function fnCategorySet(callbackval){
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/!/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$(document).find("#selectcate").empty().append("<span class='setContView'>" + catearr3[0] + "</span>");
	$("#catecode").val(catearr3[1]);
	$("#catedepth").val(catearr3[2]);
	$("#isDefault").val(catearr3[3]);
}

function fnOptionSet(callbackval){
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/!/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$("#optsetend").addClass("setContView");
	if(catearr3[0]=="editOption"){
		var arroptname = catearr3[2].replace(/,/g, "|");
		var arroptname2 = catearr3[2].replace(/,/g, "','");
		var tarroptname=eval("['" + arroptname2 + "']");
		var arroptprice = catearr3[3].replace(/,/g, "|");
		var arroptbuyprice = catearr3[4].replace(/,/g, "|");
		var arroptlimitno = catearr3[5].replace(/,/g, "|");
		var optlimitinfo = catearr3[5];
		var optlimityn=catearr3[6];
		var optlimitinfo2 = optlimitinfo.replace(/,/g, "','");
		var arroptlimitinfo=eval("['" + optlimitinfo2 + "']");
		var optionv = "";
		var optiontmp = "";
		var optlimitcnt=0;
		var optvalue = 11; // 전용옵션(11 - 99)
		for(var i = 0; i < tarroptname.length; i++) {
			optlimitcnt = optlimitcnt + Number(arroptlimitinfo[i]);
			// 전용옵션추가
			if (optvalue > 99) {
				alert("너무많은 옵션을 추가하셨습니다.");
				return false;
			}
			optiontmp = "00" + optvalue;
			optvalue = optvalue + 1;
			if(i>0){
				optionv += ("|" + optiontmp);
			}else{
				optionv += optiontmp;
			}
		}
		if(optlimitcnt > 0 && optlimityn=="Y"){
			chgodr('','','limityn','Y');
			$("#blimity").addClass("selected");
			$("#blimitn").removeClass("selected")
			$("#limitno").val(optlimitcnt);
			//$("#limitno").attr("disabled",true);
		}
		$("#itemoptioncode2").val(optionv);
		$("#itemoptioncode3").val(arroptname);
		$("#itemoptioncode4").val(arroptprice);
		$("#itemoptioncode5").val(arroptbuyprice);
		$("#itemoptioncode6").val(arroptlimitno);
		$("#opttype").val(catearr3[0]);
		$("#optionTypename1").val(catearr3[1]);
		$("#useoptionyn").val("Y");
	}else{
		$("#opttype").val(catearr3[0]);
		$("#optionTypename1").val(catearr3[1]);
		$("#optionTypename2").val(catearr3[2]);
		$("#optionTypename3").val(catearr3[3]);
		$("#optionName1").val(catearr3[4]);
		$("#optionName2").val(catearr3[5]);
		$("#optionName3").val(catearr3[6]);
		$("#optaddprice1").val(catearr3[7]);
		$("#optaddprice2").val(catearr3[8]);
		$("#optaddprice3").val(catearr3[9]);
		$("#optaddbuyprice1").val(catearr3[10]);
		$("#optaddbuyprice2").val(catearr3[11]);
		$("#optaddbuyprice3").val(catearr3[12]);
		$("#useoptionyn").val("Y");
	}
}

function fnOptionMultiItemCountReg(){
	if($('#dboptlevel').val()=="0"){
		alert("옵션 수량 변경은 옵션 항목/가격 설정 후 가능합니다.");
	}else{
		popOptionCountSet('itemid=' + $('#itemid').val() + '&limityn=' + $('#limityn').val());
	}
}

function fnMultipleStateOptionEditEnd(TotalOptLimitNo){
	//alert(TotalOptLimitNo);
	if(TotalOptLimitNo > 0){
		chgodr('LimitCnt',2,'limityn','Y');
		$("#blimity").addClass("selected");
		$("#blimitn").removeClass("selected")
		$("#limitno").val(TotalOptLimitNo);
		$("#limitno").attr("disabled",true);
		$("#optlimitset").addClass("setContView");
	}
}
var _DBOptLevel="";
function fnOptionSetEnd(msg){
	if(msg!=""){
		alert(msg);
	}else{
		$("#optsetend").addClass("setContView");
		//$("#dboptlevel").val(_DBOptLevel);
	}
}

function fnOptionEdit(){
	if($("#dboptlevel").val()!=$("#optlevel").val() && $("#dboptlevel").val()!=="0"){
		_DBOptLevel=$('#optlevel').val();
		fnAPPpopupOptionEditSet('sellcash=' + $('#sellcash').val() + '&buycash=' + $('#buycash').val() + '&dmargin=<%= npartner.FPartnerList(0).Fdiy_margin %>&mode=reset&itemid=<%=itemid%>'+'&limityn='+$('#limityn').val(),$('#optlevel').val());

	}else{
		fnAPPpopupOptionEditSet('sellcash=' + $('#sellcash').val() + '&buycash=' + $('#buycash').val() + '&dmargin=<%= npartner.FPartnerList(0).Fdiy_margin %>&mode=edit&itemid=<%=itemid%>'+'&limityn='+$('#limityn').val(),$('#optlevel').val());
	}
}

function fnSafeInfoSet(callbackval){
	var catearr = callbackval.replace(/ /g, "");
	var catearr2 = catearr.replace(/!/g, "','");
	var catearr3=eval("['" + catearr2 + "']");
	$(document).find("#safeinfotxt").empty().append("<span class='setContView'>" + catearr3[3] + "</span>");
	var safetyNum = Base64.decode(catearr3[2]);
	$("#safetyYn").val(catearr3[0]);
	$("#safetyDiv").val(catearr3[1]);
	$("#safetyNum").val(safetyNum);
}

function fnDetailPageMove(url){
	location.href=url;
}
function fnAppCallWinConfirm(){
//상품등록 콜
    if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}
	//if ($("input[name='catecode']").val() == 0){
	//	alert("[기본] 전시 카테고리를 선택해 주세요.");
	//	return;
	//}
	if (itemreg.itemname.value == ""){
		alert("상품명을 입력해 주세요.");
		$("input[name='itemname']").focus();
		return;
	}
	//상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}
	if (itemreg.itemdiv.value == ""){
		alert("상품 구분을 선택해 주세요.");
		itemreg.itemdiv.focus();
		return;
	}
	if (itemreg.itemdiv.value == "16" || itemreg.itemdiv.value == "06") {
		if(itemreg.cstodr.value == ""){
			alert("발송 구분을 선택해 주세요(즉시 발송/제작 후 발송).");
            return;
		}
	    if (itemreg.cstodr.value == "2" && itemreg.requireMakeDay.value == "") {
	        alert("제작 기간을 입력해 주세요.");
            return;
        }
		if (itemreg.requireimgchk.value == "Y" && itemreg.requireMakeEmail.value == "") {
			alert("이미지 수신 메일을 입력해 주세요.");
			// itemreg.useoptionyn.focus();
			return;
		}
	}
	if (itemreg.makername.value == ""){
		alert("제작자를 입력해 주세요.");
		itemreg.makername.focus();
		return;
	}
	if (itemreg.sourcearea.value == ""){
		alert("원산지를 입력해 주세요.");
		itemreg.sourcearea.focus();
		return;
	}
	if (itemreg.itemsource.value == ""){
		alert("재질을 입력해 주세요.");
		itemreg.itemsource.focus();
		return;
	}
	if (itemreg.itemsize.value == ""){
		alert("상품 크기를 입력해주세요.");
		itemreg.itemsize.focus();
		return;
	}
	if (itemreg.itemWeight.value == ""){
		alert("상품 무게를 입력해주세요.");
		itemreg.itemWeight.focus();
		return;
	}
	if (itemreg.itemWeight.value>2000000000){
	    alert("상품 무게는 최대 2,000,000,000 이하로 입력해주세요.");
		itemreg.itemWeight.focus();
		return;
	}
	if (itemreg.keywords.value == ""){
		alert("검색 키워드를 입력해주세요.");
		itemreg.keywords.focus();
		return;
	}
	if (itemreg.sellcash.value*1 < 400 || itemreg.sellcash.value*1 >= 20000000){
		alert("판매 가격은 400원 이상 20,000,000원 미만으로 등록 가능합니다.");
		itemreg.sellcash.focus();
		return;
	}
	if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
        alert("마일리지는 판매가보다 클 수 없습니다.");
        return;
    }
    if(itemreg.limityn.value == "Y" && itemreg.limitno.value == ""){
        alert("한정수량을 입력해주세요!");
        itemreg.limitno.focus();
        return;
    }
    if (itemreg.useoptionyn.value == "Y") {
	    if (itemreg.optlevel.value == "1") {
	    //단일옵션
    	    if (itemreg.dboptlevel.value =="") {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }
        }else if (itemreg.optlevel.value == "2") {
        //이중옵션
            if (itemreg.dboptlevel.value =="") {
                alert("추가된 옵션이 없습니다.");
                // itemreg.useoptionyn.focus();
                return;
            }
        }
	}
    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype.value=="9"){
            alert('배송 구분을 확인해주세요. [업체 조건배송] 업체가 아닙니다.');
            return;
        }
    }
    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype.value=="7")){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        return;
    }
	//배송비 설정
    if (!itemreg.deliverytype.value){
        alert('배송 설정을 선택해주세요.');
        return;
    }
	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품에 해당하는 품목을 선택해주십시요.');
        return;
    }
	//안전인증정보
    if (itemreg.safetyYn.value=="Y"){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        return;
	    }
    }
	if (itemreg.imgbasic.value == ""){
		alert("기본 이미지 한개는 등록 하셔야 합니다.");
		return;
	}
    if(confirm("상품을 수정하시겠습니까?") == true){
        if (itemreg.itemvideo.value.length>0){
            itemreg.itemvideo.value = itemreg.itemvideo.value.replace(/iframe/gi, "BUFiframe");
        }
		itemreg.action="/apps/academy/itemmaster/DIYItem_Noimg_Edit_Process_App.asp";
		itemreg.target = "FrameCKP";
        itemreg.submit();
    }
}

function fnEditItemSaveEnd(savetime){
	//alert("ok");
	$('#savetime').append(savetime);
	$('#alert3').fadeIn(800).css("display","");
	setTimeout(function(){
			$("#alert3").fadeOut(1000);
		}, 5000);
	$('#alert3').fadeIn(800).css("display","none");
}

function fnPreviewItem(){
	fnAPPpopupItemRegPreview('<%=g_AdminURL%>/apps/academy/preview/shop_prd_app.asp?itemid=<%=itemid%>');
}
var _Optiondiv;
function fnCheckResetOption(optiondiv){
	_Optiondiv=optiondiv;
	document.checkfrm.mode.value="CheckResetSingleOption";
	document.checkfrm.action="/apps/academy/itemmaster/popup/DIYItemOptionEdit_Process.asp";
	document.checkfrm.target = "FrameCKP";
	document.checkfrm.submit();
}

function fnOptionCheckEditReal(optiondiv){
	if($("#optlevel").val() != optiondiv && optiondiv!="0" && $("#optlevel").val()!="0" && $("#useoptionyn").val()=="Y"){
		fnCheckResetOption(optiondiv);
	}else{
		if(optiondiv=="0"){ 
			chgodr('setopt',1,'optlevel',0);chgodr('','','useoptionyn','N');
			$("#setoptcnt").css("display","none");
		}else if(optiondiv=="1"){
			$("#optionTypename1").val("");
			$("#optionTypename2").val("");
			$("#optionTypename3").val("");
			$("#optionName1").val("");
			$("#optionName2").val("");
			$("#optionName3").val("");
			$("#optaddprice1").val("");
			$("#optaddprice2").val("");
			$("#optaddprice3").val("");
			$("#optaddbuyprice1").val("");
			$("#optaddbuyprice2").val("");
			$("#optaddbuyprice3").val("");
			chgodr('setopt',0,'optlevel',optiondiv);
			$("#setoptcnt").css("display","");
		}else if(optiondiv=="2"){
			$("#optionTypename1").val("");
			chgodr('setopt',0,'optlevel',optiondiv);
			$("#setoptcnt").css("display","");
		}
		for (var ix = 0; ix < 3; ix++){
			if(ix==optiondiv){
				eval("$('#optbtn" + optiondiv + "')").addClass("selected");
			}else{
				eval("$('#optbtn" + ix + "')").removeClass("selected");
			}
		}
	}
}

function fnOptionDelCheckEnd(msg){
	if(msg==""){
		if(confirm("옵션 설정 변경 시\n기존 저장된 옵션이 초기화됩니다.\n변경하시겠습니까?")){
			if(_Optiondiv=="1"){
				$("#optionTypename1").val("");
				$("#optionTypename2").val("");
				$("#optionTypename3").val("");
				$("#optionName1").val("");
				$("#optionName2").val("");
				$("#optionName3").val("");
				$("#optaddprice1").val("");
				$("#optaddprice2").val("");
				$("#optaddprice3").val("");
				$("#optaddbuyprice1").val("");
				$("#optaddbuyprice2").val("");
				$("#optaddbuyprice3").val("");
				chgodr('setopt',0,'optlevel',_Optiondiv);
				$("#setoptcnt").css("display","");
			}else if(_Optiondiv=="2"){
				$("#optionTypename1").val("");
				chgodr('setopt',0,'optlevel',_Optiondiv);
				$("#setoptcnt").css("display","");
			}
			for (var ix = 0; ix < 3; ix++){
				if(ix==_Optiondiv){
					eval("$('#optbtn" + _Optiondiv + "')").addClass("selected");
				}else{
					eval("$('#optbtn" + ix + "')").removeClass("selected");
				}
			}
			
		}
	}else{
		alert(msg);
	}
}

function AddRealDetailInfo(){
	if($("#dicheckcnt").val()>14){
		alert("상세 정보의 추가 갯수는 15개 입니다.");
	}else{
		// 행추가
		var oRow;
		oRow = "							<li id='DetailList" + Number(Number($("#dicheckcnt").val())+1) + "'>"
		oRow += "								<p id='imgArea" + Number(Number($("#dicheckcnt").val())+1) + "'><button type='button' class='btnImgRegist' onclick=fnAPPuploadAddImageReal('addimgname"+ Number(Number($("#dicheckcnt").val())+1) +"','"+ Number(Number($("#dicheckcnt").val())+1) +"');>이미지 등록</button></p>"
		oRow += "								<p class='tMar1-5r'><textarea placeholder='내용을 입력해주세요' class='autosize' name='addimgtext'></textarea><input type='hidden' name='addimgname' id='addimgname'></p>"
		oRow += "							</li>"
		$("#DetailInfo ul").append(oRow);
		$("#dicheckcnt").val(Number(Number($("#dicheckcnt").val())+1));//추가 수량 카운트
		//alert($("#dicheckcnt").val());
	}
}

//-->
</script>
<script type="text/javascript" src="/apps/academy/lib/confirm.js"></script>
<script type="text/javascript" src="/apps/academy/lib/waititemreg.js"></script>
<style type="text/css">
.selectBtn2, .list li.selectBtn2 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn2 div {display:table-cell; vertical-align:middle;}
.selectBtn2 .selected {background-color:#a6d216; color:#fff;}
.selectBtn2 .grid2, .selectBtn2 .grid3, .selectBtn2 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn2 .grid2:first-child, .selectBtn2 .grid3:first-child, .selectBtn2 .grid4:first-child {padding-left:0;}
.selectBtn2 .grid2:last-child, .selectBtn2 .grid3:last-child, .selectBtn2 .grid4:last-child {padding-right:0;}

.selectBtn1, .list li.selectBtn1 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn1 div {display:table-cell; vertical-align:middle;}
.selectBtn1 .selected {background-color:#a6d216; color:#fff;}
.selectBtn1 .grid2, .selectBtn1 .grid3, .selectBtn1 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn1 .grid2:first-child, .selectBtn1 .grid3:first-child, .selectBtn1 .grid4:first-child {padding-left:0;}
.selectBtn1 .grid2:last-child, .selectBtn1 .grid3:last-child, .selectBtn1 .grid4:last-child {padding-right:0;}

.selectBtn3, .list li.selectBtn3 {display:table; width:100%; padding:1.5rem 1rem;}
.selectBtn3 div {display:table-cell; vertical-align:middle;}
.selectBtn3 .selected {background-color:#a6d216; color:#fff;}
.selectBtn3 .grid2, .selectBtn3 .grid3, .selectBtn3 .grid4 {padding-left:0.25rem; padding-right:0.25rem;}
.selectBtn3 .grid2:first-child, .selectBtn3 .grid3:first-child, .selectBtn3 .grid4:first-child {padding-left:0;}
.selectBtn3 .grid2:last-child, .selectBtn3 .grid3:last-child, .selectBtn3 .grid4:last-child {padding-right:0;}
</style>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<form name="itemreg" method="post" autocomplete="off">
		<input type="hidden" name="itemid" id="itemid" value="<%=itemid%>">
		<input type="hidden" name="imgbasic" id="imgbasic" value="<%=oitem.FOneItem.Fbasicimagecheck%>">
		<% if (oitemimg.GetImageAddByIdxIMGOnly(0,1) <> "") then %>
		<input type="hidden" name="imgadd1" id="imgadd1" value="<%=oitemimg.GetImageAddByIdxIMGOnly(0,1)%>">
		<% Else %>
		<input type="hidden" name="imgadd1" id="imgadd1">
		<% End If %>
		<% if (oitemimg.GetImageAddByIdxIMGOnly(0,2) <> "") then %>
		<input type="hidden" name="imgadd2" id="imgadd2" value="<%=oitemimg.GetImageAddByIdxIMGOnly(0,2)%>">
		<% Else %>
		<input type="hidden" name="imgadd2" id="imgadd2">
		<% End If %>
		<input type="hidden" name="itemvideo" id="itemvideo">
		<input type="hidden" name="opttype" id="opttype">
		<input type="hidden" name="optionTypename1" id="optionTypename1">
		<input type="hidden" name="optionTypename2" id="optionTypename2">
		<input type="hidden" name="optionTypename3" id="optionTypename3">
		<input type="hidden" name="optionName1" id="optionName1">
		<input type="hidden" name="optionName2" id="optionName2">
		<input type="hidden" name="optionName3" id="optionName3">
		<input type="hidden" name="optaddprice1" id="optaddprice1">
		<input type="hidden" name="optaddprice2" id="optaddprice2">
		<input type="hidden" name="optaddprice3" id="optaddprice3">
		<input type="hidden" name="optaddbuyprice1" id="optaddbuyprice1">
		<input type="hidden" name="optaddbuyprice2" id="optaddbuyprice2">
		<input type="hidden" name="optaddbuyprice3" id="optaddbuyprice3">
		<input type="hidden" name="designerid" value="<%= makerid %>">
		<input type="hidden" name="defultmargine" value="<%= npartner.FPartnerList(0).Fdiy_margin %>">
		<input type="hidden" name="defaultmaeipdiv" value="<%= npartner.FPartnerList(0).Fmaeipdiv %>">
		<input type="hidden" name="defaultFreeBeasongLimit" value="<%= npartner.FPartnerList(0).FdefaultFreeBeasongLimit %>">
		<input type="hidden" name="defaultDeliverPay" value="<%= npartner.FPartnerList(0).FdefaultDeliverPay %>">
		<input type="hidden" name="defaultDeliveryType" value="<%= npartner.FPartnerList(0).FdefaultDeliveryType %>">
		<input type="hidden" name="cd1" value="999">
		<input type="hidden" name="cd2" value="999">
		<input type="hidden" name="cd3" value="999">
		<input type="hidden" name="catecode" id="catecode">
		<input type="hidden" name="catedepth" id="catedepth">
		<input type="hidden" name="isDefault" id="isDefault">
		<input type="hidden" name="itemdiv" id="itemdiv" value="<%=oitem.FOneItem.Fitemdiv %>">
		<input type="hidden" name="cstodr" id="cstodr" value="<%=oitem.FOneItem.Fcstodr %>">
		<input type="hidden" name="reqMsg" id="reqMsg">
		<input type="hidden" name="requireimgchk" id="requireimgchk" value="<%=oitem.FOneItem.Frequireimgchk %>">
		<input type="hidden" name="vatYn" id="vatYn" value="<%=oitem.FOneItem.FvatYn %>">
		<input type="hidden" name="limityn" id="limityn" value="<%=oitem.FOneItem.Flimityn %>">
		<input type="hidden" name="useoptionyn" id="useoptionyn" value="<% If oitem.FOneItem.Foptioncnt < 1 Then %>N<% Else %>Y<% End If %>">
		<input type="hidden" name="optlevel" id="optlevel" value="<%=optionlevel%>">
		<input type="hidden" name="optwintitle" id="optwintitle">
		<input type="hidden" name="keywords" id="keywords" value="<%=oitem.FOneItem.Fkeywords %>">
		<input type="hidden" name="safetyYn" id="safetyYn" value="<%=oitem.FOneItem.FsafetyYn %>">
		<input type="hidden" name="safetyDiv" id="safetyDiv" value="<%=oitem.FOneItem.FsafetyDiv %>">
		<input type="hidden" name="safetyNum" id="safetyNum" value="<%=oitem.FOneItem.FsafetyNum %>">
		<input type="hidden" name="infoCd" id="infoCd">
		<input type="hidden" name="infoChk" id="infoChk">
		<input type="hidden" name="infoCont" id="infoCont">
		<input type="hidden" name="infoDiv" id="infoDiv" value="<%=oitem.FOneItem.FinfoDiv %>">
		<input type="hidden" name="tempSaveYn" id="tempSaveYn" value="N">
		<input type="hidden" name="deliverytype" id="deliverytype" value="<%=oitem.FOneItem.Fdeliverytype %>">
		<input type="hidden" name="dboptlevel" id="dboptlevel" value="<%=optionlevel%>">
		<input type="hidden" name="delmode" id="delmode">
		<input type="hidden" name="delfilename" id="delfilename">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">작품 등록</h1>
			<div class="artDetailInfo">
				<ul class="listTab">
					<li onclick="fnDetailPageMove('<%=g_AdminURL%>/apps/academy/itemmaster/artDetail.asp?itemid=<%=itemid%>&makerid=<%=makerid%>')"><div>기본 정보</div></li>
					<li class="current" onclick="fnDetailPageMove('<%=g_AdminURL%>/apps/academy/itemmaster/artItemEdit.asp?itemid=<%=itemid%>&makerid=<%=makerid%>')"><div>수정</div></li>
				</ul>
				<div class="artRegist tPad3r">
					<div class="registUnit"><!-- for dev msg : 비활성화 시 class : disabled 붙여주세요 -->
						<div class="basicImgRegist">
							<div class="swiper-container">
								<div class="swiper-wrapper">
									<div class="swiper-slide" id="imgspan1">
									<% If oitem.FOneItem.Fbasicimagecheck <> "" Then %>
										<img src="<%=oitem.FOneItem.Flistimage%>" onclick="fnAPPReuploadRealImage('imgbasic','basic');" />
									<% Else %>
										<button type="button" onclick="fnAPPuploadRealImage('imgbasic','basic');">이미지 등록1</button>
									<% End If %>
									</div>
									<div class="swiper-slide" id="imgspan2">
									<% if (oitemimg.GetImageAddByIdx(0,1) <> "") then %>
										<img src="<%= oitemimg.GetImageAddByIdx(0,1) %>" onclick="fnAPPReuploadRealImage('imgadd1','add1');" />
									<% Else %>
										<button type="button" onclick="fnAPPuploadRealImage('imgadd1','add1');">이미지 등록2</button>
									<% End If %>
									</div>
									<div class="swiper-slide" id="imgspan3">
									<% if (oitemimg.GetImageAddByIdx(0,2) <> "") then %>
										<img src="<%= oitemimg.GetImageAddByIdx(0,2) %>" onclick="fnAPPReuploadRealImage('imgadd2','add2');" />
									<% Else %>
										<button type="button" onclick="fnAPPuploadRealImage('imgadd2','add2');">이미지 등록3</button>
									<% End If %>
									</div>
									<div class="swiper-slide<% If videoFullUrl<> "" Then %> done<% End If %>" id="imgspan4"><button type="button" onclick="fnAPPpopupVod('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popVodAddEdit.asp?itemid=<%=itemid%>');">동영상 등록</button></div>
								</div>
							</div>
						</div>
						<ul class="list">
							<!-- li class="critical" onclick="fnAPPpopupCategory('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popCategorySelectEdit.asp?itemid=<%=itemid%>');" -->
							<li>
								<dfn><b>카테고리 설정</b></dfn>
								<div id="selectcate"><% If getCategoryInfo(itemid)<>"" Then %><span class='setContView'><%=getCategoryInfo(itemid)%></span><% End If %></div>
							</li>
							<li class="critical">
								<dfn><b>상품명</b></dfn>
								<div><input type="text" name="itemname" maxlength="64" value="<%= oitem.FOneItem.Fitemname %>" placeholder="22자 이하로 입력해주세요" id="[on,off,off,off][상품명]"/></div>
							</li>
							<li class="selectBtn">
								<div class="grid2"><button type="button" value="01" class="btnM1 btnGry<% If oitem.FOneItem.Fitemdiv="01" Then %> selected<% End If %>" onclick="chgodr('CustomOrder',1,'itemdiv','01');">일반 상품</button></div>
								<div class="grid2"><button type="button" value="06" class="btnM1 btnGry<% If oitem.FOneItem.Fitemdiv="06" Or oitem.FOneItem.Fitemdiv="16" Then %> selected<% End If %>" onclick="chgodr('CustomOrder',2,'itemdiv','16');">주문제작 상품</button></div>
							</li>
						</ul>
					</div>

					<!-- for dev msg : 주문제작 상품 선택시 노출됩니다. -->
					<div class="registUnit orderArt" id="CustomOrder" style="display:<% If oitem.FOneItem.Fitemdiv="01" Then %>none<% End If %>">
						<h2 class="critical"><b>주문제작 설정</b></h2>
						<ul class="list">
							<li class="selectBtn">
								<div class="grid2"><button type="button" onclick="chgodr('MakeDay',1,'cstodr',1);chgodr('MakeDay2',1,'','');chgodr('MakeDay3',1,'','');" class="btnM1 btnGry<% If oitem.FOneItem.Fcstodr="1" Then %> selected<% End If %>">즉시 발송</button></div>
								<div class="grid2"><button type="button" onclick="chgodr('MakeDay',2,'cstodr',2);chgodr('MakeDay2',2,'','');chgodr('MakeDay3',2,'','');" class="btnM1 btnGry<% If oitem.FOneItem.Fcstodr="2" Then %> selected<% End If %>">제작 후 발송</button></div>
							</li>
							<li class="critical" id="MakeDay" style="display:<% If oitem.FOneItem.Fcstodr="1" Then %>none<% End If %>">
								<dfn><b>제작 기간</b></dfn>
								<div><input type="number" name="requireMakeDay" maxlength="2" value="<%= oitem.FOneItem.FrequireMakeDay %>" placeholder="3" /></div>
								<div style="width:1.6rem">일</div>
							</li>
							<li class="" onclick="fnAPPpopupReqContents('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popSpecialNoteEdit.asp?itemid=<%=itemid%>');" id="MakeDay2" style="display:<% If oitem.FOneItem.Fcstodr="1" Then %>none<% End If %>">
								<dfn><b>특이사항</b><input type="hidden" id="requirecontents" name="requirecontents" value="<%=oitem.FOneItem.Frequirecontents%>"></dfn>
								<div class="listButton btnCtgySet" id="requirecontentstxt"><span class="setContView"><%=oitem.FOneItem.Frequirecontents%></span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
							</li>
							<li class="selectBtn2" id="MakeDay3" style="display:<% If oitem.FOneItem.Fcstodr="1" Then %>none<% End If %>">
								<div class="grid2"><button type="button" name="prdmsg" id="prdmsg" class="btnM1 btnGry ckBtn<% If oitem.FOneItem.Fitemdiv="06" Then %> selected<% End If %>" onclick="MultiSelectButton('prdmsg','reqMsg','06');checkItemDiv();">제작 메시지 필요</button></div>
								<div class="grid2"><button type="button" name="prdimg" id="prdimg" class="btnM1 btnGry ckBtn<% If oitem.FOneItem.Frequireimgchk="Y" Then %> selected<% End If %>" onclick="MultiSelectButton('prdimg','requireimgchk','Y');chgodr3();">제작 이미지 필요</button></div>
							</li>
							<li class="critical" onclick="#" id="MakeDay4" style="display:<% If oitem.FOneItem.Frequireimgchk<>"Y" Then %>none<% End If %>"><!-- for dev msg : 제작 이미지 필요 선택시 노출됩니다. -->
								<dfn><b>이미지 수신 메일</b></dfn>
								<div><input type="email" name="requireMakeEmail" value="<%=oitem.FOneItem.FrequireMakeEmail%>" placeholder="id1234@example.com" /></div>
							</li>
						</ul>
					</div>
					<!--// for dev msg : 주문제작 상품 선택시 노출됩니다. -->

					<div class="registUnit basicInfo">
						<h2>기본 정보</h2>
						<ul class="list">
							<li class="critical">
								<dfn><b>제작자</b></dfn>
								<div><input type="text" name="makername" maxlength="32" value="<%=oitem.FOneItem.Fmakername%>" placeholder="작가명/법인을 입력해주세요" id="[on,off,off,off][제조사]" /></div>
							</li>
							<li class="critical">
								<dfn><b>원산지</b></dfn>
								<div><input type="text" name="sourcearea" maxlength="64" value="<%=oitem.FOneItem.Fsourcearea%>" placeholder="국가명을 입력해주세요" id="[on,off,off,off][원산지]" /></div>
							</li>
							<li class="critical">
								<dfn><b>재질</b></dfn>
								<div><input type="text" name="itemsource" maxlength="64" value="<%=oitem.FOneItem.Fitemsource%>" placeholder="예) 플라스틱, 합금, 은" id="[on,off,off,off][재질]" /></div>
							</li>
							<li class="critical">
								<dfn><b>크기</b></dfn>
								<div><input type="text" name="itemsize" maxlength="64" value="<%=oitem.FOneItem.Fitemsize%>" placeholder="예) 7.5 * 7.5" id="[on,off,off,off][크기]" /></div>
								<div style="width:2.4rem">cm</div>
							</li>
							<li class="critical">
								<dfn><b>무게</b></dfn>
								<div><input type="number" name="itemWeight" maxlength="12" value="<%=oitem.FOneItem.FitemWeight%>" placeholder="예) 785" id="[on,off,off,off][무게]" pattern="[0-9]*" inputmode="numeric" min="0" /></div>
								<div style="width:1.4rem">g</div>
							</li>
							<li class="critical" onclick="fnAPPpopupKeyWord('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popKeywordEdit.asp?itemid=<%=itemid%>');">
								<dfn><b>검색 키워드</b></dfn>
								<div class="listButton btnCtgySet" id="keywordtxt"><% If oitem.FOneItem.Fkeywords <> "" Then %><span class="setContView"><%=KeyWordCnt+1%>건 등록</span><% End If %></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
							</li>
						</ul>
					</div>
					<div class="registUnit salePrice disabled" style="display:none">
						<h2 class="critical"><b>판매 가격 <span>(부가세 포함)</span></b></h2>
						<ul class="list">
							<li class="selectBtn">
								<div class="grid2"><button type="button" name="bvatYn" value="Y" class="btnM1 btnGry<% If oitem.FOneItem.FvatYn="Y" Then %> selected<% End If %>">과세</button></div>
								<div class="grid2"><button type="button" name="bvatYn" value="N" class="btnM1 btnGry<% If oitem.FOneItem.FvatYn="N" Then %> selected<% End If %>"">면세</button></div>
							</li>
							<li>
								<dfn><b>공급 마진</b></dfn>
								<div><input type="number" name="margin" maxlength="32" value="<%=fnPercent(oitem.FOneItem.Forgsuplycash,oitem.FOneItem.Forgprice,1)%>" readonly placeholder="100" /></div>
								<div style="width:1.8rem">%</div>
							</li>
							<li class="critical">
								<dfn><b>판매가</b><input type="hidden" name="sellvat"></dfn>
								<div><input type="number" name="sellcash" id="sellcash" onKeyUp="CalcuAuto(itemreg);" maxlength="7" placeholder="판매가(소비자가)를 입력해주세요"  value="<%=oitem.FOneItem.Fsellcash %>" /></div>
							</li>
							<li>
								<dfn><b>공급가</b><input type="hidden" name="buyvat"></dfn>
								<div><input type="number" name="buycash" id="buycash" maxlength="16" placeholder="0" value="<%=oitem.FOneItem.Fbuycash%>" readonly /></div>
							</li>
							<input type="hidden" name="mwdiv" value="<%=oitem.FOneItem.Fmwdiv%>"> <!-- 매입위탁구분 :업체배송 -->
							<input type="hidden" name="sellyn" value="<%=oitem.FOneItem.Fsellyn%>">
							<input type="hidden" name="isusing" value="<%=oitem.FOneItem.Fisusing%>">
							<input type="hidden" name="mileage" value="<%=oitem.FOneItem.Fmileage%>">
						</ul>
					</div>
					<div class="registUnit quantity">
						<h2 class="critical"><b>수량 설정</b></h2>
						<ul class="list">
							<li class="selectBtn3">
								<div class="grid2"><button type="button" class="btnM1 btnGry<% If oitem.FOneItem.Flimityn="Y" Then %> selected<% End If %>" name="blimityn" id="limitbtn1" value="Y" onclick="fnLimitCheckOption('Y');">한정 수량</button></div>
								<div class="grid2"><button type="button" class="btnM1 btnGry<% If oitem.FOneItem.Flimityn="N" Then %> selected<% End If %>" name="blimityn" id="limitbtn2" value="N" onclick="fnLimitCheckOption('N');">무제한</button></div>
							</li>
							<li id="LimitCnt" style="display:<% If oitem.FOneItem.Flimityn="N" Then %>none<% ElseIf oitemoption.FResultCount > 1 Then %>none<% End If %>"><!--for dev msg : 한정수량 선택시 노출됩니다. -->
								<dfn><b>수량</b></dfn>
								<div><input type="number" name="limitno" id="limitno" value="<%=oitem.FOneItem.Flimitno%>" placeholder="수량을 입력해주세요" /></div>
								<div style="width:1.6rem">개</div>
							</li>
						</ul>
					</div>
					<div class="registUnit option">
						<h2 class="critical"><b>옵션 설정</b></h2>
						<ul class="list">
							<li class="selectBtn1">
								<div class="grid3"><button type="button" name="boptlevel" id="optbtn0" value="0" class="btnM1 btnGry<% If oitemoption.FResultCount<1 Then %> selected<% End If %>" onClick="fnOptionCheckEditReal(0);">사용안함</button></div>
								<div class="grid3"><button type="button" name="boptlevel" id="optbtn1" value="1" class="btnM1 btnGry<% If oitemoption.FResultCount > 1 And oitemoption.IsMultipleOption=false Then %> selected<% End If %>" onClick="fnOptionCheckEditReal(1);">단일 옵션</button></div>
								<div class="grid3"><button type="button" name="boptlevel" id="optbtn2" value="2" class="btnM1 btnGry<% If oitemoption.IsMultipleOption Then %> selected<% End If %>" onClick="fnOptionCheckEditReal(2);">이중 옵션</button></div>
							</li>
							<li class="critical" id="setopt" onclick="fnOptionEdit();" style="display:<% If oitemoption.FResultCount<1 Then %>none<% End If %>"><!--for dev msg : 단일 옵션 or 이중 옵션 선택시 노출됩니다. -->
								<dfn><b>항목/가격</b></dfn>
								<div class="listButton btnCtgySet"><span id="optsetend" class="<% If oitemoption.FResultCount > 1 Then %>setContView<% End If %>">설정됨</span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
							</li>
							<li class="" id="setoptcnt" onclick="fnOptionMultiItemCountReg();" style="display:<% If oitemoption.FResultCount<1 Then %>none<% End If %>">
								<dfn><b>수량</b></dfn>
								<div class="listButton btnCtgySet"><span id="optlimitset" class="<% If optlimit.FTotalCount > 0 Then %>setContView<% End If %>">설정됨</span></div>
							</li>
						</ul>
					</div>
					<div class="registUnit delivery">
						<h2 class="critical"><b>배송 설정 <span>(부가세 포함)</span></b></h2>
						<ul class="list">
							<li class="selectBtn">
								<div class="grid4"><button type="button" name="bdeliverytype" value="2" class="btnM1 btnGry<% If oitem.FOneItem.Fdeliverytype="2" Then %> selected<% End If %>" onClick="chgodr('',1,'deliverytype',2);">무료</button></div>
								<div class="grid4"><button type="button" name="bdeliverytype" value="9" class="btnM1 btnGry<% If oitem.FOneItem.Fdeliverytype="9" Then %> selected<% End If %>"" onClick="chgodr('',1,'deliverytype',9);">조건부</button></div>
								<div class="grid4"><button type="button" name="bdeliverytype" value="7" class="btnM1 btnGry<% If oitem.FOneItem.Fdeliverytype="7" Then %> selected<% End If %>"" onClick="chgodr('',1,'deliverytype',7);">착불</button></div>
							</li>
							<li class="" onclick="fnAPPpopupDeliveryInfo('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popDeliveryInfoEdit.asp?itemid=<%=itemid%>');">
								<dfn><b>배송비 안내</b><input type="hidden" id="ordercomment"  name="ordercomment" value="<%=oitem.FOneItem.Fordercomment%>"></dfn>
								<div class="listButton btnCtgySet" id="deliveryInfotxt"><span class="setContView"><%=oitem.FOneItem.Fordercomment%></span></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
							</li>
							<li style="display:none">
								<dfn><b>교환 / 환불 정책</b></dfn>
								<div><p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize" name="refundpolicy">- 반품/환불은 상품수령일로부터 7일 이내만 가능합니다. 
- 출고 이후 환불요청 시 상품 회수 후 처리됩니다. 
- 변심 반품의 경우 왕복배송비를 차감한 금액이 환불되며, 제품 및 포장 상태가 재판매 가능하여야 합니다. 
- 상품 불량인 경우는 배송비를 포함한 전액이 환불됩니다.
- 완제품으로 수입된 상품의 경우 A/S가 불가합니다. 
- 교환/환불/배송비안내/AS에 대한 개별기준이 상품페이지에 있는 경우 작가님의 개별기준이 우선 적용 됩니다.</textarea></p></div><!-- for dev msg : 디폴트는 span태그 display:none 상태입니다. 설정하고 난 후 class="setContView" 붙여주세요 -->
							</li>
						</ul>
					</div>
					<div class="registUnit law">
						<h2 class="critical"><b>관련법 필수 입력 항목</b></h2>
						<ul class="list">
							<li class="critical" onclick="fnAPPpopupItemInfoDiv('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popArtInfoEdit.asp?itemid=<%=oitem.FOneItem.Fitemid%>');">
								<dfn><b>상품정보제공고시</b></dfn>
								<div class="listButton btnCtgySet" id="iteminfotxt"><span class="setContView"><%=oitem.FOneItem.getinfoDivName%></span></div>
							</li>
							<li class="critical" onclick="fnAPPpopupSafeInfo('<%=g_AdminURL%>/apps/academy/itemmaster/popup/popArtSafeEdit.asp?itemid=<%=oitem.FOneItem.Fitemid%>');">
								<dfn><b>안전인증대상</b></dfn>
								<div class="listButton btnCtgySet" id="safeinfotxt"><span class="setContView"><%=oitem.FOneItem.getsafetyDivName%></span></div>
							</li>
						</ul>
					</div>
<%
dim oaddimg
set oaddimg = new CItemAddImage
oaddimg.FRectItemID = itemid
oaddimg.GetItemAddImageList
%>
					<div class="detail" id="DetailInfo">
						<div class="registUnit">
							<h2 class="critical"><b>상세 정보</b><input type="hidden" name="dicheckcnt" id="dicheckcnt" value="<% If oaddimg.FResultCount < 1 Then %>1<% Else %><%=oaddimg.FResultCount%><% End If %>"></h2>
							<ul class="list">
								<% If oaddimg.FResultCount>0 Then %>
								<% For i=0 To oaddimg.FResultCount - 1 %>
								<li id="DetailList<%=i+1%>">
									<p id="imgArea<%=i+1%>"><% If oaddimg.FITemList(i).FADDIMAGEName="" Then %><button type="button" class="btnImgRegist" onclick="fnAPPuploadAddImageReal('addimgname<%=i+1%>','DetailList<%=i+1%>');">이미지 등록</button><% Else %><img src="<%=oaddimg.GetImageAddByIdx(2,i+1)%>" alt="" onclick="fnAPPReuploadAddImageReal('addimgname<%=i+1%>','<%=i+1%>');" /><% End If %><input type="hidden" name="addimgname" id="addimgname" value="<%=oaddimg.FITemList(i).FADDIMAGEName%>"></p>
									<p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize" name="addimgtext"><%=oaddimg.FITemList(i).FADDIMGTXT%></textarea></p>
								</li>
								<% Next %>
								<% Else %>
								<li id="DetailList1">
									<p id="imgArea1"><button type="button" class="btnImgRegist" onclick="fnAPPuploadAddImageReal('addimgname1','1');">이미지 등록</button><input type="hidden" name="addimgname" id="addimgname"></p>
									<p class="tMar1-5r"><textarea placeholder="내용을 입력해주세요" class="autosize" name="addimgtext"></textarea></p>
								</li>
								<% End If %>
							</ul>
						</div>
						<div class="addBtn">
							<button type="button" class="btnB1 btnDkGry" id="addbtn" onClick="AddRealDetailInfo()"><span class="itemAdd">추가</span></button>
							<p class="tPad2r">최대 15개까지 등록 가능합니다.</p>
						</div>
					</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<!-- 알림 메세지 -->
		<!-- 알림 메세지 -->
		<div class="attentionBar" style="display:none" id="alert3">
			<p><img src="http://image.thefingers.co.kr/apps/2016/blt_save.png" alt="저장표시" style="width:1.2rem; height:1.2rem; margin:0.3rem 0.3rem 0 0" /> <span id="savetime"></span>에 저장되었습니다.</p>
		</div>
		<!-- 하단 플로팅 버튼 -->
		<div class="floatingBar">
			<p><button type="button" class="btnV16a btnRed2V16a" onClick="fnPreviewItem();">미리보기</button></p>
		</div>
		</form>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<form method="post" name="checkfrm">
	<input type="hidden" name="mode">
	<input type="hidden" name="itemid" id="itemid" value="<%=itemid%>">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</body>
</html>
<%
Set npartner = Nothing
Set oitem = Nothing
Set oitemoption = Nothing
Set oitemimg = Nothing
Set oaddimg = Nothing
Set optlimit = Nothing
Set ovod = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->