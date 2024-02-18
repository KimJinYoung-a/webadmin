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
Dim waititemid, limityn
waititemid = requestCheckVar(request("waititemid"),10)
limityn = requestCheckVar(request("limityn"),1)
If limityn="" Then limityn="N"

if waititemid="" then
	Response.Write "<script>alert('잘못된 접속입니다. (파라메터)');fnAPPclosePopup();</script>"
	Response.end
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = waititemid
oitemoption.GetWaitItemMultiOptionInfo

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
	fnAPPShowRightConfirmBtns();
});


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

function fnAppCallWinConfirm(){
	//alert("ok");
	document.soption.action="/apps/academy/itemmaster/waitItemMultipleStateOptionEdit.asp";
	document.soption.target="FrameCKP";
	document.soption.submit();
}
function fnOptionStateEditEnd(TotalOptLimitNo){
	fnAPPopenerJsCallClose("fnMultipleStateOptionEditEnd(\"" + TotalOptLimitNo + "\")");
}

</script>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">이중 옵션 설정</h1>
			<form method="post" name="soption" autocomplete="off">
			<input type="hidden" name="waititemid" id="waititemid" value="<%=waititemid%>">
			<div class="registUnit optSet tMar2r" style="display:<% If oitemoption.FResultCount<1 Then %>none<% End If %>">
				<h2><b>옵션별 수량 설정</b></h2>
				<ul class="list">
					<% If oitemoption.FTotalMultipleNo >0 Then %>
					<% For i=0 To oitemoption.FResultCount - 1 %>
					<li class="">
						<dfn><em><%= oitemoption.FITemList(i).Fitemoption %></em><b><%= oitemoption.FITemList(i).FOptionName %></b></dfn>
						<% If limityn="Y" Then %>
						<div><input type="number" name="optlimitno" id="optlimitno" placeholder="0" value="<%= oitemoption.FITemList(i).Foptlimitno %>" /><input type='hidden' name='optionyn' value='<%= oitemoption.FITemList(i).Foptisusing %>'><input type='hidden' name='itemoption' value='<%= oitemoption.FITemList(i).Fitemoption %>'></div>
						<div style="width:1.5rem">개</div>
						<% Else %>
						<div class="rt">무제한<input type='hidden' name='optionyn' value='<%= oitemoption.FITemList(i).Foptisusing %>'><input type='hidden' name='itemoption' value='<%= oitemoption.FITemList(i).Fitemoption %>'><input type='hidden' name='optlimitno' value='<%= oitemoption.FITemList(i).Foptlimitno %>'></div>
						<% End If %>
						<div class="lPad3r">
							<span class="btnSetting">
								<label>옵션 사용여부</label>
								<button type="button" class="<% if oitemoption.FITemList(i).Foptisusing = "Y" Then %>settingOn<% End If %>" onclick="fnOptionEachDisable('<%=i%>')" id="eachopt<%=i%>">옵션 사용여부 설정</button>
							</span>
						</div>
					</li>
					<% Next %>
					<% End If %>
				</ul>
			</div>
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
set oitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->