<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 단일 옵션 설정"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<%
Dim OptionDiv, WinTitle, ItemDefaultMargin, SellCash, BuyCash, DefaultMargin, thisUrl, waititemid, limityn
OptionDiv = request("OptionDiv")
SellCash = request("sellcash")
BuyCash = request("buycash")
DefaultMargin = request("dmargin")
waititemid = request("waititemid")
limityn = requestCheckVar(request("limityn"),1)
If limityn="" Then limityn="N"

If DefaultMargin="" Then
	ItemDefaultMargin = 100-CLng(BuyCash/SellCash*100*100)/100
Else
	ItemDefaultMargin = DefaultMargin
End If

dim oitemoption, optionTypename
set oitemoption = new CItemOption
oitemoption.FRectItemID = waititemid
if waititemid<>"" then
	oitemoption.GetWaitItemOptionInfo
	If oitemoption.FResultCount > 0 Then 
		optionTypename=oitemoption.FITemList(0).FoptionTypename
	Else
		optionTypename=""
	End If
Else
	optionTypename=""
end if


Dim i
thisUrl=request.ServerVariables("PATH_INFO") + "?" + request.ServerVariables("QUERY_STRING")
%>
<script>
function fnReSetOption(){
	if(confirm("옵션을 초기화 하시겠습니까?")){
		$("#mode").val("editOptionDel");
		document.frmOpt.action="/apps/academy/itemmaster/popup/WaitDIYItemPopupDetailinfoEdit_Process.asp";
		document.frmOpt.target="FrameCKP";
		document.frmOpt.submit();
	}
}

function fnDetailInfoEnd2(){
	location.replace("<%=thisUrl%>");
}

var VItemDefaultMargin = <%= ItemDefaultMargin %>;

function AutoCalcuBuyPrice(j){
	if (!$("#optdiv" + j + " input[name=optaddbuyprice]").length){
        $("#optdiv" + j + " input[name=optaddprice]").val(parseInt($("#optdiv" + j + " input[name=optaddprice]").val()*1*(100-VItemDefaultMargin)/100));
    }else{
        $("#optdiv" + j + " input[name=optaddbuyprice]").val(parseInt($("#optdiv" + j + " input[name=optaddprice]").val()*1*(100-VItemDefaultMargin)/100));
    }
}

function fnAppCallWinConfirm(){
	var catename;
	var arrcode;
	var arrdepth;

    var frm = document.frmOpt;
    var optAddpriceExists = false;
	var optTypeNm1='', optTypeNm2='', optTypeNm3='';
	var arroptNm1='', arroptaddprice1='', arroptaddbuyprice1='';
	var arroptNm2='', arroptaddprice2='', arroptaddbuyprice2='';
	var arroptNm3='', arroptaddprice3='', arroptaddbuyprice3='';
	var loop=0;
	var param,formcheck;
	formcheck=true;

	$("#Opt input[name='optionName']").each(function(i){
		var str=$("#Opt input[name='optionName']:eq(" + i + ")").val();
		var limit = 96; //제한byte를 가져온다.
		var strLength = 0;
		var strTitle = "";
		var strPiece = "";
		var check = false;
		for (ix = 0; ix < str.length; ix++){
			var code = str.charCodeAt(ix);
			var ch = str.substr(ix,1).toUpperCase();
			//체크 하는 문자를 저장
			strPiece = str.substr(ix,1)
			code = parseInt(code);
			if ((ch < "0" || ch > "9") && (ch < "A" || ch > "Z") && ((code > 255) || (code < 0))){
				strLength = strLength + 3; //UTF-8 3byte 로 계산
			}else{
				strLength = strLength + 1;
			}
			if(strLength>limit){ //제한 길이 확인
				check = true;
				break;
			}else{
				strTitle = strTitle+strPiece; //제한길이 보다 작으면 자른 문자를 붙여준다.
			}
		}
		if(check){
			$("#Opt input[name='optionName']:eq(" + i + ")").val(strTitle);
			alert("옵션 내 항목 길이가 "+limit+"byte 초과 되었습니다. 초과된 문자는 잘려서 입력 됩니다.");
		}
	});


	//단일옵션
	if (frm.optionTypename.value.length<1){
		alert('옵션 구분명을 입력하세요.');
		frm.optionTypename.focus();
		formcheck=false;
					return false;
	}
	if (!frm.optionName.length){
		if (frm.optionName.value.length<1){
			alert('옵션명을 입력하세요.');
			frm.optionName.focus();
			formcheck=false;
					return false;
		}
	}else{
		for (var i=0;i<frm.optionName.length;i++){
			if (frm.optionName[i].value.length<1){
				alert('옵션명을 입력하세요.');
				frm.optionName[i].focus();
				formcheck=false;
					return false;
			}
			
			//옵션명이 중복되는지 체크.
			for (var j=0;j<frm.optionName.length;j++){
				if ((i!=j)&&(frm.optionName[i].value==frm.optionName[j].value)){
					alert('옵션명을 중복하여 사용할 수 없습니다. - [' + frm.optionName[j].value + ']');
					frm.optionName[j].focus();
					formcheck=false;
					return false;
				}
			}
			
		}
	}
	
	//추가금액
	if (!frm.optaddprice.length){
		if (frm.optaddprice.value.length<1){
			alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
			frm.optaddprice.focus();
			formcheck=false;
					return false;
		}
		
		if (!IsDigit(frm.optaddprice.value)){
			alert('추가금액은 숫자만 가능합니다.');
			frm.optaddprice.focus();
			formcheck=false;
					return false;
		}
		
		if ((frm.optaddbuyprice.value*1)>(frm.optaddprice.value*1)) {
			alert('공급가가 매입가 보다 클 수 없습니다.');
			frm.optaddbuyprice.focus();
			formcheck=false;
					return false;
		}
		
		if ((frm.optaddprice.value*1>0) && (frm.optaddbuyprice.value*1!=parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100))) {
			if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
				frm.optaddbuyprice.focus();
				formcheck=false;
					return false;
			}
		}
		
		optAddpriceExists = (optAddpriceExists||(frm.optaddprice.value*1>0));
	}else{
		for (var i=0;i<frm.optaddprice.length;i++){
			if (frm.optaddprice[i].value.length<1){
				alert('추가금액을 입력하세요. (추가금액이 없으면 0)');
				frm.optaddprice[i].focus();
				formcheck=false;
					return false;
			}
			
			if (!IsDigit(frm.optaddprice[i].value)){
				alert('추가금액은 숫자만 가능합니다.');
				frm.optaddprice[i].focus();
				formcheck=false;
					return false;
			}
			
			if ((frm.optaddbuyprice[i].value*1)>(frm.optaddprice[i].value*1)) {
				alert('공급가가 매입가 보다 클 수 없습니다.');
				frm.optaddbuyprice[i].focus();
				formcheck=false;
					return false;
			}
			
			if ((frm.optaddprice[i].value*1>0) && (frm.optaddbuyprice[i].value*1!=parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100))) {
				if (!confirm('옵션 추가 금액에 대한 공급 금액이 상품 기본 마진 (<%= ItemDefaultMargin %>) 공급액(' + parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100) + '원) 과 일치 하지 않습니다. 계속 하시겠습니까?')){
					frm.optaddbuyprice[i].focus();
					formcheck=false;
					return false;
				}
			}
			
			optAddpriceExists = (optAddpriceExists||(frm.optaddprice[i].value*1>0));
		}
	}
	
	//추가금액-공급가
	if (!frm.optaddbuyprice.length){
		if (frm.optaddbuyprice.value.length<1){
			alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
			frm.optaddbuyprice.focus();
			formcheck=false;
					return false;
		}
		
		if (!IsDigit(frm.optaddbuyprice.value)){
			alert('추가금액 공급가는 숫자만 가능합니다.');
			frm.optaddbuyprice.focus();
			formcheck=false;
					return false;
		}
	}else{
		for (var i=0;i<frm.optaddbuyprice.length;i++){
			if (frm.optaddbuyprice[i].value.length<1){
				alert('추가금액 공급가를 입력하세요. (추가금액이 없으면 0)');
				frm.optaddbuyprice[i].focus();
				formcheck=false;
					return false;
			}
			
			if (!IsDigit(frm.optaddbuyprice[i].value)){
				alert('추가금액 공급가는 숫자만 가능합니다.');
				frm.optaddbuyprice[i].focus();
				formcheck=false;
					return false;
			}
		}
	}
	if ($("input[name=optionName]").length < 2) {
		alert("옵션은 두개 이상이어야 합니다.(옵션별로 한정/전시설정이 가능합니다.)");
		formcheck=false;
		return false;
	}
	if(formcheck){
		var optionv = "";
		var optiontmp = "";
		var optvalue = 11; // 전용옵션(11 - 99)
		for(var i = 0; i < $("input[name=optionName]").length; i++) {
			// 전용옵션추가
			if (optvalue > 99) {
				alert("너무많은 옵션을 추가하셨습니다.");
				return false;
			}
			optiontmp = "00" + optvalue;
			optvalue = optvalue + 1;
			if(i>0){
				optionv += ("," + optiontmp);
			}else{
				optionv += optiontmp;
			}
		}
		$("#optioncode").val(optionv);
		document.frmOpt.action="/apps/academy/itemmaster/popup/WaitDIYItemPopupDetailinfoEdit_Process.asp";
		document.frmOpt.target="FrameCKP";
		document.frmOpt.submit();
	}

}

function fnDetailInfoEnd(){
	var jSonTXT;
	var arrOptName, arrOptAddPrice, arrOptBuyPrice;
	arrOptName = new Array();
	arrOptAddPrice = new Array();
	arrOptBuyPrice = new Array();
	$("input[name=optionName]").each(function(idx){
		 arrOptName[idx] = $("input[name=optionName]:eq(" + idx + ")").val();
		 arrOptAddPrice[idx] = $("input[name=optaddprice]:eq(" + idx + ")").val();
		 arrOptBuyPrice[idx] = $("input[name=optaddbuyprice]:eq(" + idx + ")").val();
	});
	jSonTXT = JSON.stringify({"mode":$("#mode").val(), "optiontypename":$("input[name=optionTypename]").val(), "optionname":arrOptName, "optionaddprice":arrOptAddPrice, "optionbuyprice":arrOptBuyPrice, "optioncode":$("#optioncode").val(),"itemid":<%=waititemid%>});
	jSonTXT = Base64.encode(jSonTXT);
	//alert(jSonTXT);
	fnAPPopenerJsCallClose("fnOptionSet(\"" + jSonTXT + "\")");
}

$(function() {

	//$("input[name='optionName']").keyup(function(){
		fnAPPShowRightConfirmBtns();
	//});

	//옵션 추가
	$("#Opt #addbtn").click(function(){
		// 기존에 값에 중복 옵션 여부 검사
		if($("#checkopt").val()!=0){
			$("input[name='optNm_temp']").val($("#Opt input[name='optionName']:eq(" + ($("#Opt input[name='optionName']").length-1) + ")").val());//중복체크용 현재 폼 값
		}
		var CheckOverlap=false;
		$("#Opt input[name='optionName']").each(function(i){
			if(i < $("#Opt input[name='optionName']").length-1){
				if($("#Opt input[name='optionName']:eq(" + i + ")").val() == $("input[name='optNm_temp']").val()){
					//alert($("input[name='optionName']:eq(" + i + ")").val() + "/"+$("input[name='optionName_temp']").val());
					CheckOverlap=true;
					return false;
				}
				else{
					CheckOverlap=false;
					return;
				}
			}
		});

		if(CheckOverlap){
			alert("같은 이름의 옵션이 있습니다.");
		}else if($("#checkopt").val()>20){
			alert("단일 옵션의 추가 갯수는 20개 입니다.");
		}else{
			// 행추가
			var oRow;
			oRow = "							<dd id='optdiv" + (Number($("#checkopt").val())+1) + "'>"
			oRow += "								<ul>"
			oRow += "									<li>"
			oRow += "										<div>"
			oRow += "											<span><input type='text' name='optionName' id='optionName' placeholder='옵션 내 항목' style='width:100%;' /></span>"
			oRow += "										</div>"
			oRow += "									</li>"
			oRow += "									<li>"
			oRow += "										<div>"
			oRow += "											<span><input type='number' name='optaddprice' id='optaddprice' placeholder='추가금액' style='width:100%;' onKeyUp='AutoCalcuBuyPrice(" + (Number($("#checkopt").val())+1) + ");' /></span>"
			oRow += "											<span>원</span>"
			oRow += "										</div>"
			oRow += "									</li>"
			oRow += "									<li>"
			oRow += "										<div>"
			oRow += "											<span><input type='number' name='optaddbuyprice' id='optaddbuyprice' placeholder='공급가' style='width:100%;' readonly /></span>"
			oRow += "											<span>원</span>"
			oRow += "										</div>"
			oRow += "									</li>"
			oRow += "								</ul>"
			oRow += "								<button type='button' class='btnM1 btnGry tMar1r' onclick='fnOptionEachDisable(" + (Number($("#checkopt").val())+1) + ")'>옵션삭제</button>"
			oRow += "							</dd>"
			$("#Opt dl").append(oRow);
			$("#checkopt").val(Number($("#checkopt").val())+1);//옵션 추가 수량 카운트
			$(".optionUnit button").css("display","");
			$("#Opt #addbtn").attr("disabled",true);//버튼 비활성화
			//alert($("#checkopt").val());
		}
	});

	//옵션 폼 체크 후 추가버튼 활성화(단일 옵션)
	$("#Opt").keyup(function(){
		var formcheck=true;
		if(formcheck){
			//alert($("#Opt input[name='optionName']:eq(0)").val());
			$("#Opt input[name='optionName']").each(function(x){
				//alert($("#Opt input[name='optionName']:eq(" + x + ")").val());
				if($("#Opt input[name='optionName']:eq(" + x + ")").val() == ""){
					formcheck=false;
					return false;
				}
				else{
					formcheck=true;
					return;
				}
			});
		}
		if(formcheck){
			$("#Opt #addbtn").attr("disabled",false);
		}
	});
<% if optionTypename <> "" then %>
	$("#Opt #addbtn").attr("disabled",false);
	fnAPPShowRightConfirmBtns();
<% end if %>

	$('#optionTypename').blur(function(){
		var thisObject = $(this);
		var str=thisObject.val();
		var limit = 32; //제한byte를 가져온다.
		var strLength = 0;
		var strTitle = "";
		var strPiece = "";
		var check = false;
		for (i = 0; i < str.length; i++){
			var code = str.charCodeAt(i);
			var ch = str.substr(i,1).toUpperCase();
			//체크 하는 문자를 저장
			strPiece = str.substr(i,1)
			code = parseInt(code);
			if ((ch < "0" || ch > "9") && (ch < "A" || ch > "Z") && ((code > 255) || (code < 0))){
				strLength = strLength + 3; //UTF-8 3byte 로 계산
			}else{
				strLength = strLength + 1;
			}
			if(strLength>limit){ //제한 길이 확인
				check = true;
				break;
			}else{
				strTitle = strTitle+strPiece; //제한길이 보다 작으면 자른 문자를 붙여준다.
			}
		}
		$('#optionTypename').val(strTitle);
		if(check){
			alert("옵션이름 길이가 "+limit+"byte 초과되었습니다. 초과된 문자는 잘려서 입력 됩니다.");
		}
	});

});

function fnOptionEachDisable(num){
	if($("#Opt input[name='optionName']").length<=2){
		$(".optionUnit button").css("display","none");
	}
	$("#optdiv"+num).remove();
	$("#Opt #addbtn").attr("disabled",false);
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<form name="frmOpt" method="post" onsubmit="return false;" autocomplete="off">
		<input type="hidden" name="mode" id="mode" value="editOption">
		<input type="hidden" name="waititemid" value="<%=waititemid%>">
		<input type='hidden' name='optNm_temp' value='' />
		<input type="hidden" name="checkopt" id="checkopt" value="<% If oitemoption.FResultCount >0 Then Response.write oitemoption.FResultCount-1 Else Response.write "0" End If %>">
		<input type="hidden" name="useoptionyn" value="Y">
		<input type="hidden" name="optlevel" value="1">
		<input type="hidden" name="optioncode" id="optioncode">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">옵션 설정</h1>
			<div class="optionSetting<% If limityn="N" Then %> multiOpt<% End If %>">
				<div class="optSetListWrap" id="Opt">
					<div class="setList">
						<dl class="optionUnit">
							<dt>
								<div><input type="text" name="optionTypename" id="optionTypename" value="<%=optionTypename%>" placeholder="옵션이름" style="width:100%;" /></div>
							</dt>
							<% If oitemoption.FResultCount>0 Then %>
							<% For i=0 To oitemoption.FResultCount - 1 %>
							<dd id="optdiv<%=i%>">
								<ul>
									<li>
										<div>
											<span><input type="text" name="optionName" value="<%= oitemoption.FITemList(i).FOptionName %>" id="optionName" placeholder="옵션 내 항목" style="width:100%;" /></span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" name="optaddprice" id="optaddprice" value="<%= oitemoption.FITemList(i).Foptaddprice %>" placeholder="추가금액" style="width:100%;" onKeyUp="AutoCalcuBuyPrice(<%=i%>);" /></span>
											<span>원</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" name="optaddbuyprice" id="optaddbuyprice" value="<%= oitemoption.FITemList(i).Foptaddbuyprice %>" placeholder="공급가" style="width:100%;" readonly /></span>
											<span>원</span>
										</div>
									</li>
								</ul>
								<button type="button" class="btnM1 btnGry tMar1r" onclick="fnOptionEachDisable('<%=i%>')"<% If oitemoption.FResultCount<2 Then %> style="display:none"<% End If %>>옵션삭제</button>
							</dd>
							<% Next %>
							<% Else %>
							<dd id="optdiv0">
								<ul>
									<li>
										<div>
											<span><input type="text" name="optionName" id="optionName" placeholder="옵션 내 항목" style="width:100%;" /></span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" name="optaddprice" id="optaddprice" placeholder="추가금액" style="width:100%;" onKeyUp="AutoCalcuBuyPrice(0);" /></span>
											<span>원</span>
										</div>
									</li>
									<li>
										<div>
											<span><input type="number" name="optaddbuyprice" id="optaddbuyprice" placeholder="공급가" style="width:100%;" readonly /></span>
											<span>원</span>
										</div>
									</li>
								</ul>
								<button type="button" class="btnM1 btnGry tMar1r" onclick="fnOptionEachDisable('0')" style="display:none">옵션삭제</button>
							</dd>
							<% End If %>
						</dl>
					</div>
					<div class="addBtn">
						<button type="button" class="btnB1 btnDkGry" disabled="disabled" id="addbtn"><span class="itemAdd">추가</span></button><!-- for dev msg : 추가버튼 클릭시 setList division의 optionUnit dd가 추가(최대 9개까지)되면 됩니다.-->
					</div>
				</div>
			</div>
		</div>
		<!--// content -->
		<!-- 하단 플로팅 버튼 -->
		</form>
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->