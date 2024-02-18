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
<%
Dim itemid, ItemDefaultMargin, SellCash, BuyCash, DefaultMargin, OptionDiv, mode
OptionDiv = requestCheckVar(request("OptionDiv"),1)
SellCash = requestCheckVar(request("sellcash"),10)
BuyCash = requestCheckVar(request("buycash"),10)
DefaultMargin = requestCheckVar(request("dmargin"),5)
itemid = requestCheckVar(request("itemid"),10)
mode = request("mode")

If DefaultMargin="" Then
	ItemDefaultMargin = 100-CLng(BuyCash/SellCash*100*100)/100
Else
	ItemDefaultMargin = DefaultMargin
End If

dim coitemoption, OptionCount
OptionCount="0"
set coitemoption = new CItemOption
coitemoption.FRectItemID = itemid
if itemid<>"" then
	coitemoption.GetItemOptionCountInfo
	If coitemoption.FResultCount > 0 Then 
		OptionCount=coitemoption.FResultCount
	Else
		OptionCount="0"
	End If
end If

Dim i
%>
<script>
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
<% if OptionCount <> "" then %>
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
	$("input[name='multioptcnt']").val(Number(Number($("input[name='multioptcnt']").val())+1));
	fnAPPShowRightConfirmBtns();
}

function fnAppCallWinConfirm(){
	//저장된 작품은 바로 수정
	if($("input[name='multioptcnt']").val()<2){
		alert("이중옵션을 사용할 경우 옵션구분명 은 최소 2개 이상 등록하셔야 합니다.");
		formcheck=false;
		return false;
	}
	if($("#changeyn").val()=="Y"){
		document.moption.action="/apps/academy/itemmaster/popup/DIYItemOptionEdit_Process.asp";
		document.moption.target="FrameCKP";
		document.moption.submit();
		$.showLoading({name: 'circle-turn',allowHide: true});
	}else{
		var param = "N";
		fnAPPopenerJsCallClose("fnOptionNoEditSet(\"" + param + "\")");
	}
}

function fnOptionEachDisable(num){
	//alert(num + "/" + $("input[name='optionyn']:eq(" + num + ")").val());
	if ($("input[name='optionyn']:eq(" + num + ")").val()=="N"){
		$("input[name='optionyn']:eq(" + num + ")").val("Y");
		$("#eachopt"+num).removeClass();
	}else{
		$("#eachopt"+num).addClass("settingOn");
		$("input[name='optionyn']:eq(" + num + ")").val("N");
	}
}

function fnMultipleOptionEditEnd(msg){
	fnAPPopenerJsCallClose("fnOptionSetEnd(\"" + msg + "\")");
}

function fnCheckOptionInput(querystring,OptionDiv){
	if(OptionDiv=="2" && $("#optionTypename1").val()==""){
		alert("옵션은 순차적으로 입력하시기 바랍니다.");
	}else if(OptionDiv=="3" && $("#optionTypename2").val()==""){
		alert("옵션은 순차적으로 입력하시기 바랍니다.");
	}else{
		fnAPPpopupMultiOptionEdit(querystring);
	}
}

</script>
<script src="/apps/academy/lib/jquery.loading.min.js"></script>
<link href="/apps/academy/lib/loading.min.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
Dim optionTypename1, optionTypename2, optionTypename3
dim oitemoptionM
set oitemoptionM = new CItemOption
oitemoptionM.FRectItemID = itemid
if itemid<>"" then
	oitemoptionM.GetItemMultipleOptionInfo
end If
%>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">이중 옵션 설정</h1>
			<form method="post" name="moption" autocomplete="off" id="fopt">
			<% If mode="edit" Then %>
			<input type="hidden" name="mode" id="mode" value="addDoubleOption">
			<% Else %>
			<input type="hidden" name="mode" id="mode" value="ResetMultiOptionCustom">
			<% End If %>
			<input type="hidden" name="itemid" id="itemid" value="<%=itemid%>">
			<input type="hidden" name="changeyn" id="changeyn" value="N">
<%
If oitemoptionM.FResultCount>0 Then
	For i=0 To oitemoptionM.FResultCount - 1
		If oitemoptionM.FItemList(i).FTypeSeq="1" Then
			optionTypename1=oitemoptionM.FItemList(i).FoptionTypename
			Response.write "<input type='hidden' name='optname1' id='optname1' value='" + CStr(oitemoptionM.FItemList(i).FoptionKindName) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optaddprice1' id='optaddprice1' value='" + CStr(oitemoptionM.FItemList(i).Foptaddprice) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optbuyprice1' id='optbuyprice1' value='" + CStr(oitemoptionM.FItemList(i).Foptaddbuyprice) + "'>" + vbcrlf
		ElseIf oitemoptionM.FItemList(i).FTypeSeq="2" Then
			optionTypename2=oitemoptionM.FItemList(i).FoptionTypename
			Response.write "<input type='hidden' name='optname2' id='optname2' value='" + CStr(oitemoptionM.FItemList(i).FoptionKindName) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optaddprice2' id='optaddprice2' value='" + CStr(oitemoptionM.FItemList(i).Foptaddprice) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optbuyprice2' id='optbuyprice2' value='" + CStr(oitemoptionM.FItemList(i).Foptaddbuyprice) + "'>" + vbcrlf
		Else
			optionTypename3=oitemoptionM.FItemList(i).FoptionTypename
			Response.write "<input type='hidden' name='optname3' id='optname3' value='" + CStr(oitemoptionM.FItemList(i).FoptionKindName) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optaddprice3' id='optaddprice3' value='" + CStr(oitemoptionM.FItemList(i).Foptaddprice) + "'>" + vbcrlf
			Response.write "<input type='hidden' name='optbuyprice3' id='optbuyprice3' value='" + CStr(oitemoptionM.FItemList(i).Foptaddbuyprice) + "'>" + vbcrlf
		End If		
	Next
End If
%>
			<input type="hidden" name="opttypename1" id="opttypename1" value="<%=optionTypename1%>">
			<input type="hidden" name="opttypename2" id="opttypename2" value="<%=optionTypename2%>">
			<input type="hidden" name="opttypename3" id="opttypename3" value="<%=optionTypename3%>">
			<input type="hidden" name="multioptcnt" value="<% =OptionCount %>">
			<input type="hidden" name="limityn" id="limityn" value="N">
			<input type="hidden" name="limitno" id="limitno" value="0">
			<ul class="list">
				<li class="" onClick="fnCheckOptionInput('OptionDiv=1&sellcash=<%=SellCash%>&buycash=<%=BuyCash%>&dmargin=<%=DefaultMargin%>&itemid=<%=itemid%>','1');">
					<dfn><b>옵션 1</b></dfn>
					<div class="listButton btnCtgySet" id="opt1"><span class='setContView'><%=optionTypename1%></span></div>
				</li>
				<li class="" onClick="fnCheckOptionInput('OptionDiv=2&sellcash=<%=SellCash%>&buycash=<%=BuyCash%>&dmargin=<%=DefaultMargin%>&itemid=<%=itemid%>','2');">
					<dfn><b>옵션 2</b></dfn>
					<div class="listButton btnCtgySet" id="opt2"><span class='setContView'><%=optionTypename2%></span></div>
				</li>
				<li class="" onClick="fnCheckOptionInput('OptionDiv=3&sellcash=<%=SellCash%>&buycash=<%=BuyCash%>&dmargin=<%=DefaultMargin%>&itemid=<%=itemid%>','3');">
					<dfn><b>옵션 3</b></dfn>
					<div class="listButton btnCtgySet" id="opt3"><span class='setContView'><%=optionTypename3%></span></div>
				</li>
			</ul>
			</form>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
set oitemoptionM = Nothing
set coitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->