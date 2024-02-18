<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

If not (Request.ServerVariables("REMOTE_ADDR") = "61.252.133.75" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.105" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.106") Then
	Response.End
End If
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
Dim i, cCurator, vIdx, vUnitArr, vContents
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cCurator = New CSearchMng
	cCurator.FRectIdx = vIdx
	cCurator.FRectOnlyUnitList = "o"
	cCurator.sbCuratorDetail

	vUnitArr = cCurator.FUnitArr

	Set cCurator = Nothing
Else
	Response.Write "<script>alert('�߸��� �����Դϴ�.');windows.close();</script>"
	dbget.close
	Response.End
End If
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
.popWinV17 {overflow:hidden; position:absolute; left:0; top:0; right:0; bottom:0; width:100%; height:100%; font-family:"malgun Gothic","�������", Dotum, "����", sans-serif;}
.popWinV17 h1 {height:40px; padding:12px 15px 0; color:#fff; font-size:17px; background:#c80a0a; border-bottom:1px solid #d80a0a}
.popWinV17 h2 {position:relative; padding:12px 15px; color:#333; font-size:12px; font-weight: bold; background-color:#444; border-top:1px solid #666; font-family:"malgun Gothic","�������", Dotum, "����", sans-serif; z-index:55; color:#fff;}
.popContainerV17 {position:absolute; left:0; top:40px; right:0; bottom:90px; width:100%; border-bottom:1px solid #ddd;}
.contL {position:relative; width:57%; height:100%; border-right:1px solid #ddd; z-index:10; overflow-y:auto;}
.contR {position:absolute; right:0; top:0; bottom:0; width:38%; height:100%; border-left:1px solid #ddd;}
.tbListWrap {position:relative; width:100%; height:100%;}
.tbDataList, .thDataList {display:table; width:100%;}
.tbDataList li, .thDataList li {display:table; width:100%; margin-top:-1px; border-top:1px solid #ddd; border-bottom:1px solid #ddd; }
.thDataList li {height:33px; background-color:#eaeaea; border-top:2px solid #ccc; font-weight:bold;}
.tbDataList li {background-color:#fff; z-index:100;}
.tbDataList li p, .thDataList li p {display:table-cell; padding:7px; text-align:center; vertical-align:middle; line-height:1.4;}
.thDataList li p {white-space:nowrap;}
.handling {background-color:rgba(42,42,57,0.2) !important; height:30px; border:none;}
#sortable li {cursor:move;}
.popBtnWrap {position:absolute; left:0; bottom:0; width:100%; height:60px; text-align:center;}
.textOverflow {width:100%; display:block; text-overflow:ellipsis; overflow:hidden; white-space:nowrap;}
.btnMove {position:absolute; left:59.5%; top:50%; width:40px; height:70px; margin-top:-35px; margin-left:-20px; padding:0; border:none; background:transparent url(/images/btn_move_arrow.png) no-repeat 50% 50%; z-index:1000; cursor:pointer;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
$(function(){
	jsCallUnitlist("item");
	$("#isfirstloading").val("x");
	
	
	//���ָ���Ʈ ���콺 �巹�� ����
	$("#sortable").sortable({
		placeholder:"handling",
		start: function(event, ui) {
		},
		stop: function(){
			var i=9999;
			$(this).parent().find("input[name^='sort']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='sort']").each(function(){
				$(this).val(i);
				i++;
			});
			jsSortReSetting();
		}
	}).disableSelection();
});


function jsSortReSetting(){
	var cnt = $("input[name=sort]").length;
	var newc = "";
	for(var n=0; n<cnt; n++){
		newc = newc + $("input[name=svalue]").eq(n).val() + ",";
	}
	$("#contents").val(newc);
	
	jsUnitListSetting();
}


//����¡
function NextPage(ipage,g){
	$("#page").val(ipage);
	//if ((document.frm.itemname.value.length>0)&&(ipage*1==1)){
	//    alert('��ǰ�� �˻��� ����� �ִ� 1000���� ���ѵ˴ϴ�.');  // 2������ fulltext �˻��� ���ι������ ����.
	//}

	jsCallUnitlist(g);
}


//������ ����Ʈ ��������
function jsCallUnitlist(g){
	var url;
	if(g == "event"){
		url = "keywordQratingUnitEventlistAjax.asp";
	}else{
		url = "keywordQratingUnitItemlistAjax.asp";
	}

	if(g == "event"){
		$("#unitTypeitem").empty();
	}else{
		$("#unitTypeevent").empty();
	}

	if($("#isfirstloading").val() == "x"){
		
		$("#btnsearh1").hide();
		$("#btnsearh2").show();


		var formData = $("#"+g+"frm").serialize().replace(/=(.[^&]*)/g,
			function($0,$1){ 
			return "="+escape(decodeURIComponent($1)).replace('%26','&').replace('%3D','=')
		});
		

		$.ajax({
				url: "/admin/search/temp/"+url+"",
				type: "GET",
				cache: false,
				data: formData,
				success: function(message)
				{
					$("#unitType"+g+"").empty().append(message);
				}
		});
	}else{
		$.ajax({
				url: "/admin/search/temp/"+url+"",
				cache: false,
				success: function(message)
				{
					$("#unitType"+g+"").empty().append(message);
				}
		});
	}

	$(".unitPannel").hide();
	$("#unitType"+g+"").show();
	$(".tab > ul > li").removeClass("selected");
	$("#btn"+g+"").addClass("selected");
	$("#nowunit").val(g);
}


//������ ����Ʈ üũ�ڽ� ���� �κ� Ŭ����
function jsThisClick(i,g){
	if($("#contentsidx"+i+"").is(":checked")){
		$("#contentsidx"+i+"").prop("checked", false);
	}else{
		if(jsIsUnitExist(i,g)){
			//�����ʿ� �ִ� ���. stop.
			alert("�̹� ���õ� Unit �Դϴ�.");
			return;
		}else{
			//�����ʿ� ���� ���. �״�� ����.
			$("#contentsidx"+i+"").prop("checked", true);
		}
	}
	$("#tr"+i+"").css('backgroundColor', '#D9FFFF');
	jsThisCheck(i,g);
}


//������ ����Ʈ üũ�ڽ� Ŭ��
function jsThisCheck(i,g){
	if($("#contentsidx"+i+"").is(":checked")){
		if(jsIsUnitExist(i,g)){
			//�����ʿ� �ִ� ���. stop.
			$("#contentsidx"+i+"").prop("checked", false);
			alert("�̹� ���õ� Unit �Դϴ�.");
			return;
		}else{
			//�����ʿ� ���� ���. �״�� ����.
			$("#tr"+i+"").css('backgroundColor', '#D9FFFF');
		}
	}else{
		$("#tr"+i+"").css('backgroundColor', '#FFFFFF');
	}
	
	jsUnitSetting(i,g,'');
}


//�����ʿ� ���õ� ������ �ִ��� üũ
function jsIsUnitExist(i,g){
	var result = false;
	var newval = g + "$" + i + ",";
	var conval = $("#contents").val();
	
	if(conval.indexOf(newval) > -1){
		result = true;
	}else{
		result = false;
	}
	return result
}


//�̺�Ʈ �˻�. �˻��� üũ.
function jsEventValueCheck(){
	var frm = document.eventfrm
	if (frm.selDate.value == "A"){
			frm.iSD.value = "";
			frm.iED.value = "";
			frm.sEtxt.value = "";
	}

	if(frm.selEvt.value== "evt_code"&&frm.sEtxt.value!=""){
		frm.sEtxt.value = frm.sEtxt.value.replace(/\s/g, "");
		if(!IsDigit(frm.sEtxt.value)){
			alert("�̺�Ʈ�ڵ�� ���ڸ� �����մϴ�.");
			frm.sEtxt.value = "";
			return;
		}
	}
}


//unit ����Ʈ ����� ���� �� ����. ������ �ְ� ������ ����.
function jsUnitSetting(i,g,a){
	var newval = g + "$" + i + ",";
	var conval = $("#contents").val();
	
	if(conval.indexOf(newval) > -1){
		if(a != "setting"){
			var tmp = conval.replace(newval,"");
			$("#contents").val(tmp);
		}
	}else{
		$("#contents").val(conval+newval);
	}
}


//unit üũ�Ѱ� ���. >>��ư
function jsUnitListSetting(){
	var conval = $("#contents").val();

	$.ajax({
			url: "/admin/search/temp/keywordQratingUnitAjax.asp?idx=<%=vIdx%>&contents="+conval+"",
			cache: false,
			success: function(msgu)
			{
				if(msgu == "10"){
					alert("Unit�� 10�� �̸����� �������ּ���.");
				}else{
					$("#sortable").empty().append(msgu);
					jsLeftDelete();
					opener.location.reload();
				}
			}
	});
}


//unit ����Ʈ ����
function jsUnitDelete(g,i){
	$("#unitgubun").val(g);
	$("#unitcontentsidx").val(i);
	frm2.submit();
}


//���� ����Ʈ ����
function jsLeftDelete(){
	var g = $("#nowunit").val();
	$("input[type='checkbox']:checked").each(function() {
		var cvalue = $(this).val();
		$("#contentsidx"+cvalue+"").prop("checked", false);
		$("#tr"+cvalue+"").css('backgroundColor', '#FFFFFF');
		jsUnitSetting(cvalue,g,'setting');
	});
}


//�θ�â, �ڽ�â ��� reload
function jsAllReload(){
	opener.location.reload();
	location.reload();
}
</script>
</head>
<body>
<div class="popWinV17">
	<h1>Unit �˻�</h1>
	<input type="hidden" name="isfirstloading" id="isfirstloading" value="">
	<input type="hidden" name="nowunit" id="nowunit" value="item">
	<div class="popContainerV17">
		<div class="contL">
			<h2>Unit ����</h2>
			<div class="tab" style="margin:-1px 0 0 -1px;">
				<ul>
					<li id="btnitem" class="col11 selected" onClick="jsCallUnitlist('item');return false;">��ǰ</li>
					<li id="btnevent" class="col11 " onClick="jsCallUnitlist('event');return false;">�̺�Ʈ</li>
				</ul>
			</div>
			<!-- ��ǰ Tab -->
			<div id="unitTypeitem" class="unitPannel">
			</div>
			<!-- �̺�Ʈ Tab -->
			<div id="unitTypeevent" class="unitPannel" style="display:none;">
			</div>
		</div>

		<input type="button" class="btnMove" title="�����ؼ� ���" onClick="jsUnitListSetting();" />

		<div class="contR">
			<h2 style="margin-left:-1px;">Unit ���� ����</h2>
			<div class="tbListWrap">
				<ul class="thDataList">
					<li>
						<p class="cell15 lt"> ����</p>
						<p>Unit��</p>
						<p class="cell05"></p>
					</li>
				</ul>
				<ul id="sortable" class="tbDataList">
				<%
				If IsArray(vUnitArr) Then
					For i =0 To UBound(vUnitArr,2)
						If i = 0 Then
							vContents = vUnitArr(1,i) & "$" & vUnitArr(2,i)
						Else
							vContents = vContents & "," & vUnitArr(1,i) & "$" & vUnitArr(2,i)
						End If
				%>
						<li>
							<p class="cell15 lt"><%=vUnitArr(1,i)%></p>
							<p class="lt">
								<span class="textOverflow">
									<% If vUnitArr(1,i) = "event" AND date() > vUnitArr(4,i) Then Response.Write "<font color=red>[����]</font> " End If %>
									<%=db2html(vUnitArr(0,i))%>
								</span>
							</p>
							<p class="cell05"><input type="button" class="btn" value="����" onClick="jsUnitDelete('<%=vUnitArr(1,i)%>','<%=vUnitArr(2,i)%>');" /></p>
							<input type="hidden" id="sort" name="sort" value="<%=vUnitArr(3,i)%>">
							<input type="hidden" id="svalue" name="svalue" value="<%=vUnitArr(1,i)&"$"&vUnitArr(2,i)%>">
						</li>
				<%
					Next
					vContents = vContents & ","
				End IF
				%>
				</ul>
				<input type="hidden" id="contents" name="contents" value="<%=vContents%>" size="80">
			</div>
		</div>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="���ÿϷ�" onclick="window.close();" class="cRd1" style="width:100px; height:30px;" />
		<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
	</div>
</div>
<form name="frm2" action="keywordQratingProc.asp" method="post" target="iframeproc" style="margin:0px;">
<input type="hidden" id="action" name="action" value="unitdeletepop">
<input type="hidden" id="idx" name="idx" value="<%=vIdx%>">
<input type="hidden" id="unitgubun" name="unitgubun" value="">
<input type="hidden" id="unitcontentsidx" name="unitcontentsidx" value="">
</form>
<iframe src="about:blank" name="iframeproc" width="0" height="0" frameborder="0"></iframe>
</body>
</html>