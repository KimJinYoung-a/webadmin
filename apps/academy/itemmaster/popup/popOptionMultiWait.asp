<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Session.CodePage = 65001
Response.Charset = "UTF-8"
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 이중 옵션 설정"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/DIYitemCls.asp"-->
<!-- #include virtual="/lib/util/base64Lib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<script language="jscript" runat="server">
function jsDecodeURIComponent(v) { return decodeURIComponent(v); }
function jsEncodeURIComponent(v) { return encodeURIComponent(v); }
</script>
<%
Dim ItemDefaultMargin, SellCash, BuyCash, DefaultMargin, thisUrl, waititemid
SellCash = request("sellcash")
BuyCash = request("buycash")
DefaultMargin = request("dmargin")
waititemid = request("waititemid")

If DefaultMargin="" Then
	ItemDefaultMargin = 100-CLng(BuyCash/SellCash*100*100)/100
Else
	ItemDefaultMargin = DefaultMargin
End If

Dim param
param = Base64Decode(jsDecodeURIComponent(request("param")),"UTF-8")

Dim ArrCnt, ArrCnt1, ArrCnt2, ArrCnt3, OptionTypeName1, OptionTypeName2, OptionTypeName3, OptionName1, OptionName2, OptionName3
Dim OptionPrice1, OptionPrice2, OptionPrice3, OptionBuyPrice1, OptionBuyPrice2, OptionBuyPrice3, ix

dim jsonParse
set jsonParse = JSON.parse(param)
redim OptionName1(jsonParse.optionname1.length-1)
redim OptionPrice1(jsonParse.optionaddprice1.length-1)
redim OptionBuyPrice1(jsonParse.optionbuyprice1.length-1)
redim OptionName2(jsonParse.optionname2.length-1)
redim OptionPrice2(jsonParse.optionaddprice2.length-1)
redim OptionBuyPrice2(jsonParse.optionbuyprice2.length-1)
redim OptionName3(jsonParse.optionname3.length-1)
redim OptionPrice3(jsonParse.optionaddprice3.length-1)
redim OptionBuyPrice3(jsonParse.optionbuyprice3.length-1)
OptionTypeName1 = jsonParse.optiontypename1
OptionTypeName2 = jsonParse.optiontypename2
OptionTypeName3 = jsonParse.optiontypename3
ArrCnt1 = jsonParse.optionname1.length-1
ArrCnt2 = jsonParse.optionname2.length-1
ArrCnt3 = jsonParse.optionname3.length-1
For ix = 0 To ArrCnt1
	OptionName1(ix) = jsonParse.optionname1.get(ix)
	OptionPrice1(ix) = jsonParse.optionaddprice1.get(ix)
	OptionBuyPrice1(ix) = jsonParse.optionbuyprice1.get(ix)
Next
For ix = 0 To ArrCnt2
	OptionName2(ix) = jsonParse.optionname2.get(ix)
	OptionPrice2(ix) = jsonParse.optionaddprice2.get(ix)
	OptionBuyPrice2(ix) = jsonParse.optionbuyprice2.get(ix)
Next
For ix = 0 To ArrCnt3
	OptionName3(ix) = jsonParse.optionname3.get(ix)
	OptionPrice3(ix) = jsonParse.optionaddprice3.get(ix)
	OptionBuyPrice3(ix) = jsonParse.optionbuyprice3.get(ix)
Next
set jsonParse = Nothing

ArrCnt=0
If OptionTypeName1<>"" Then
	ArrCnt=ArrCnt+1
End If
If OptionTypeName2 <> "" Then
	ArrCnt=ArrCnt+1
End If
If OptionTypeName3 <> "" Then
	ArrCnt=ArrCnt+1
End If
'Response.write param
'Response.end
thisUrl=request.ServerVariables("PATH_INFO") + "?" + "sellcash=" + SellCash + "&buycash=" + BuyCash + "&dmargin=" + DefaultMargin
%>
<script>
function fnReSetOption(){
	location.replace("<%=thisUrl%>");
}

$(function() {
	//push setting
	$(".btnSetting button").click(function(){
		$(this).toggleClass('settingOn');
	});

	//option tab control
	function optSize() {
		var optLength = $('.optionTab li:visible').length;
		$('.optionTab li').each(function(){
			if (optLength == 1) {
				$(this).children('button').hide();
				$('.btnPlus').show();
			} else if (optLength == 2) {
				$(this).children('button').show();
				$('.btnPlus').show();
			} else if (optLength == 3) {
				$(this).children('button').show();
				$('.btnPlus').hide();
			}
		});
	}
	optSize();

	$('.optionTab li').click(function(){
		$('.optionTab li').removeClass('current');
		$(this).addClass('current');
	});

	$('.optionTab li button').click(function(e){
		e.preventDefault();
		$(this).parent('li').hide();
		optSize();
	});
<% if ArrCnt >= 0 then %>
	fnAPPShowRightConfirmBtns();
<% end if %>
});

function fnMultiOptionSet(callbackval){
	callbackval = Base64.decode(callbackval);
	var jSonTXT = JSON.parse(callbackval);
	//alert(callbackval);
	$("#opt"+jSonTXT.optiondiv).empty();
	$("#opt"+jSonTXT.optiondiv).append("<span class='setContView'>" + jSonTXT.optiontypename + "</span>");
	$("#changeyn").val(jSonTXT.changeyn);
	
	if(jSonTXT.optiondiv=="1"){
		//alert(jSonTXT.optiontypename);
		$("#opttypename1").val(jSonTXT.optiontypename);
		$("input[name=optname1]").each(function(idx){
			$("#optname1").remove();
			$("#optaddprice1").remove();
			$("#optbuyprice1").remove();
		});
		for(var i=0; i < jSonTXT.optionname.length; i++){
			$('#fopt').append('<input type="hidden" id="optname1" name="optname1" value="' + jSonTXT.optionname[i] + '">');
			$('#fopt').append('<input type="hidden" id="optaddprice1" name="optaddprice1" value="' + jSonTXT.optionaddprice[i] + '">');
			$('#fopt').append('<input type="hidden" id="optbuyprice1" name="optbuyprice1" value="' + jSonTXT.optionbuyprice[i] + '">');
		}
	}else if(jSonTXT.optiondiv=="2"){
		$("#opttypename2").val(jSonTXT.optiontypename);
		$("input[name=optname2]").each(function(idx){
			$("#optname2").remove();
			$("#optaddprice2").remove();
			$("#optbuyprice2").remove();
		});
		for(var i=0; i < jSonTXT.optionname.length; i++){
			$('#fopt').append('<input type="hidden" id="optname2" name="optname2" value="' + jSonTXT.optionname[i] + '">');
			$('#fopt').append('<input type="hidden" id="optaddprice2" name="optaddprice2" value="' + jSonTXT.optionaddprice[i] + '">');
			$('#fopt').append('<input type="hidden" id="optbuyprice2" name="optbuyprice2" value="' + jSonTXT.optionbuyprice[i] + '">');
		}
	}else{
		$("#opttypename3").val(jSonTXT.optiontypename);
		$("input[name=optname3]").each(function(idx){
			$("#optname3").remove();
			$("#optaddprice3").remove();
			$("#optbuyprice3").remove();
		});
		for(var i=0; i < jSonTXT.optionname.length; i++){
			$('#fopt').append('<input type="hidden" id="optname3" name="optname3" value="' + jSonTXT.optionname[i] + '">');
			$('#fopt').append('<input type="hidden" id="optaddprice3" name="optaddprice3" value="' + jSonTXT.optionaddprice[i] + '">');
			$('#fopt').append('<input type="hidden" id="optbuyprice3" name="optbuyprice3" value="' + jSonTXT.optionbuyprice[i] + '">');
		}
	}
	//alert($("input[name=optname1]").length);
	$("input[name='multioptcnt']").val(Number(Number($("input[name='multioptcnt']").val())+1));
	fnAPPShowRightConfirmBtns();
}

function fnAppCallWinConfirm(){
	//저장된 작품은 바로 수정
	if($("input[name='multioptcnt']").val()<2){
		alert("이중옵션을 사용할 경우 옵션구분명 은 최소 2개 이상 등록하셔야 합니다.");
		return false;
	}
	if($("#changeyn").val()=="Y"){
		document.moption.action="/apps/academy/itemmaster/waitItemMultipleOptionEdit.asp";
		document.moption.target="FrameCKP";
		document.moption.submit();
	}else{
		var param = "N";
		fnAPPopenerJsCallClose("fnOptionNoEditSet(\"" + param + "\")");
	}
}

function fnMultipleOptionEditEnd(waititemid){
	var jSonTXT;
	var arrOptName1, arrOptAddPrice1, arrOptBuyPrice1;
	var arrOptName2, arrOptAddPrice2, arrOptBuyPrice2;
	var arrOptName3, arrOptAddPrice3, arrOptBuyPrice3;
	arrOptName1 = new Array();
	arrOptAddPrice1 = new Array();
	arrOptBuyPrice1 = new Array();
	arrOptName2 = new Array();
	arrOptAddPrice2 = new Array();
	arrOptBuyPrice2 = new Array();
	arrOptName3 = new Array();
	arrOptAddPrice3 = new Array();
	arrOptBuyPrice3 = new Array();

	if($("input[name=opttypename1]").val()!=""){
		$("input[name=optname1]").each(function(idx){
			 arrOptName1[idx] = $("input[name=optname1]:eq(" + idx + ")").val();
			 arrOptAddPrice1[idx] = $("input[name=optaddprice1]:eq(" + idx + ")").val();
			 arrOptBuyPrice1[idx] = $("input[name=optbuyprice1]:eq(" + idx + ")").val();
		});
	}
	if($("input[name=opttypename2]").val()!=""){
		$("input[name=optname2]").each(function(idx){
			 arrOptName2[idx] = $("input[name=optname2]:eq(" + idx + ")").val();
			 arrOptAddPrice2[idx] = $("input[name=optaddprice2]:eq(" + idx + ")").val();
			 arrOptBuyPrice2[idx] = $("input[name=optbuyprice2]:eq(" + idx + ")").val();
		});
	}
	if($("input[name=opttypename3]").val()!=""){
		$("input[name=optname3]").each(function(idx){
			 arrOptName3[idx] = $("input[name=optname3]:eq(" + idx + ")").val();
			 arrOptAddPrice3[idx] = $("input[name=optaddprice3]:eq(" + idx + ")").val();
			 arrOptBuyPrice3[idx] = $("input[name=optbuyprice3]:eq(" + idx + ")").val();
		});
	}
	jSonTXT = JSON.stringify({"mode":$("#mode").val(), "waititemid":waititemid, "optiontypename1":$("input[name=opttypename1]").val(), "optionname1":arrOptName1, "optionaddprice1":arrOptAddPrice1, "optionbuyprice1":arrOptBuyPrice1, "optiontypename2":$("input[name=opttypename2]").val(), "optionname2":arrOptName2, "optionaddprice2":arrOptAddPrice2, "optionbuyprice2":arrOptBuyPrice2, "optiontypename3":$("input[name=opttypename3]").val(), "optionname3":arrOptName3, "optionaddprice3":arrOptAddPrice3, "optionbuyprice3":arrOptBuyPrice3});
	jSonTXT = Base64.encode(jSonTXT);
	fnAPPopenerJsCallClose("fnOptionSet(\"" + jSonTXT + "\")");
}

function fnOptionSetOpen(OptionDiv){
	if(OptionDiv=="2" && $("#opttypename1").val()==""){
		alert("옵션은 순차적으로 입력하시기 바랍니다.");
	}else if(OptionDiv=="3" && $("#opttypename2").val()==""){
		alert("옵션은 순차적으로 입력하시기 바랍니다.");
	}else{
		var jSonTXT;
		var arrOptName, arrOptAddPrice, arrOptBuyPrice;
		arrOptName = new Array();
		arrOptAddPrice = new Array();
		arrOptBuyPrice = new Array();
		if(OptionDiv=="1"){
			//alert($("input[name=optname1]").length);
			$("input[name=optname1]").each(function(idx){
				 arrOptName[idx] = $("input[name=optname1]:eq(" + idx + ")").val();
				 arrOptAddPrice[idx] = $("input[name=optaddprice1]:eq(" + idx + ")").val();
				 arrOptBuyPrice[idx] = $("input[name=optbuyprice1]:eq(" + idx + ")").val();
			});
		}else if(OptionDiv=="2"){
			$("input[name=optname2]").each(function(idx){
				 arrOptName[idx] = $("input[name=optname2]:eq(" + idx + ")").val();
				 arrOptAddPrice[idx] = $("input[name=optaddprice2]:eq(" + idx + ")").val();
				 arrOptBuyPrice[idx] = $("input[name=optbuyprice2]:eq(" + idx + ")").val();
			});
		}else{
			$("input[name=optname3]").each(function(idx){
				 arrOptName[idx] = $("input[name=optname3]:eq(" + idx + ")").val();
				 arrOptAddPrice[idx] = $("input[name=optaddprice3]:eq(" + idx + ")").val();
				 arrOptBuyPrice[idx] = $("input[name=optbuyprice3]:eq(" + idx + ")").val();
			});
		}
		jSonTXT = JSON.stringify({"opttypename":$("#opttypename"+OptionDiv).val(), "optionname":arrOptName, "optionaddprice":arrOptAddPrice, "optionbuyprice":arrOptBuyPrice});
		jSonTXT = encodeURIComponent(Base64.encode(jSonTXT));
		fnAPPpopupMultiOptionWait("OptionDiv=" + OptionDiv + "&sellcash=<%=SellCash%>&buycash=<%=BuyCash%>&dmargin=<%=DefaultMargin%>&param="+jSonTXT);
	}
}
</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">이중 옵션 설정</h1>
			<form method="post" name="moption" autocomplete="off" id="fopt">
			<input type="hidden" name="mode" id="mode" value="editOptionMultiple">
			<% If ArrCnt >= 0 Then %>
			<input type="hidden" name="changeyn" id="changeyn" value="N">
			<input type="hidden" name="waititemid" value="<%=waititemid%>">
			<input type="hidden" name="designerid" value="<%= request.cookies("partner")("userid") %>">
			<input type="hidden" name="multioptcnt" value="<% =ArrCnt %>">
			<input type="hidden" name="opttypename1" id="opttypename1" value="<%=OptionTypeName1%>">
			<input type="hidden" name="opttypename2" id="opttypename2" value="<%=OptionTypeName2%>">
			<input type="hidden" name="opttypename3" id="opttypename3" value="<%=OptionTypeName3%>">
			<input type="hidden" name="limityn" id="limityn" value="N">
			<input type="hidden" name="limitno" id="limitno" value="0">
			<% For ix = 0 To ArrCnt1 %>
			<input type="hidden" name="optname1" id="optname1" value="<%=OptionName1(ix)%>">
			<input type="hidden" name="optaddprice1" id="optaddprice1" value="<%=OptionPrice1(ix)%>">
			<input type="hidden" name="optbuyprice1" id="optbuyprice1" value="<%=OptionBuyPrice1(ix)%>">
			<% Next %>
			<% For ix = 0 To ArrCnt2 %>
			<input type="hidden" name="optname2" id="optname2" value="<%=OptionName2(ix)%>">
			<input type="hidden" name="optaddprice2" id="optaddprice2" value="<%=OptionPrice2(ix)%>">
			<input type="hidden" name="optbuyprice2" id="optbuyprice2" value="<%=OptionBuyPrice2(ix)%>">
			<% Next %>
			<% For ix = 0 To ArrCnt3 %>
			<input type="hidden" name="optname3" id="optname3" value="<%=OptionName3(ix)%>">
			<input type="hidden" name="optaddprice3" id="optaddprice3" value="<%=OptionPrice3(ix)%>">
			<input type="hidden" name="optbuyprice3" id="optbuyprice3" value="<%=OptionBuyPrice3(ix)%>">
			<% Next %>
			<ul class="list">
				<li class="" onClick="fnOptionSetOpen('1');">
					<dfn><b>옵션 1</b></dfn>
					<div class="listButton btnCtgySet" id="opt1"><span class='setContView'><%=OptionTypeName1%></span></div>
				</li>
				<li class="" onClick="fnOptionSetOpen('2');">
					<dfn><b>옵션 2</b></dfn>
					<div class="listButton btnCtgySet" id="opt2"><span class='setContView'><%=OptionTypeName2%></span></div>
				</li>
				<li class="" onClick="fnOptionSetOpen('3');">
					<dfn><b>옵션 3</b></dfn>
					<div class="listButton btnCtgySet" id="opt3"><span class='setContView'><%=OptionTypeName3%></span></div>
				</li>
			</ul>
			<% Else %>
			<input type="hidden" name="changeyn" id="changeyn" value="Y">
			<input type="hidden" name="waititemid" value="<%=waititemid%>">
			<input type="hidden" name="designerid" value="<%= request.cookies("partner")("userid") %>">
			<input type="hidden" name="multioptcnt" value="0">
			<input type="hidden" name="opttypename1" id="opttypename1">
			<input type="hidden" name="opttypename2" id="opttypename2">
			<input type="hidden" name="opttypename3" id="opttypename3">
			<input type="hidden" name="limityn" id="limityn" value="N">
			<input type="hidden" name="limitno" id="limitno" value="0">
			<ul class="list">
				<li class="" onClick="fnOptionSetOpen('1');">
					<dfn><b>옵션 1</b></dfn>
					<div class="listButton btnCtgySet" id="opt1"></div>
				</li>
				<li class="" onClick="fnOptionSetOpen('2');">
					<dfn><b>옵션 2</b></dfn>
					<div class="listButton btnCtgySet" id="opt2"></div>
				</li>
				<li class="" onClick="fnOptionSetOpen('3');">
					<dfn><b>옵션 3</b></dfn>
					<div class="listButton btnCtgySet" id="opt3"></div>
				</li>
			</ul>
			<% End If %>
			</form>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->