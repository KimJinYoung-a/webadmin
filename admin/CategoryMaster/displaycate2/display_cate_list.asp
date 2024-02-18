<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/categorymaster/displaycate2/classes/displaycateCls.asp"-->

<%
	Dim cDisp, vWidth, vDepth, vCateCode, vCateName, vUseYN, vSortNo, vCurrpage
	vCurrpage 	= NullFillWith(Request("cpg"), "1")
	vDepth 		= NullFillWith(Request("depth_s"), "1")
	vCateCode 	= Request("catecode_s")
	vCateName	= Request("catename_s")
	vUseYN 		= Request("useyn_s")
	vSortNo 	= Request("sortno_s")
	
	
	Dim vNotCateReg, makerid, cdl, cdm, cds, itemid_s, itemname, keyword, sellyn, usingyn, danjongyn, limityn, sailyn, deliverytype, sortDiv, pagesize
	vNotCateReg	= Request("notcatereg")
	makerid		= request("makerid")
	cdl 		= request("cdl")
	cdm 		= request("cdm")
	cds 		= request("cds")
	itemid_s	= request("itemid_s")
	itemname	= request("itemname")
	keyword		= request("keyword")
	sellyn      = request("sellyn")
	usingyn     = request("usingyn")
	danjongyn   = request("danjongyn") 
	limityn     = request("limityn") 
	sailyn      = request("sailyn")
	deliverytype = request("deliverytype")
	sortDiv		= request("sortDiv")
	pagesize	= request("pagesize")
	

	'// 상품배열 재지정
	if itemid_s<>"" then
		dim iA ,arrTemp,arrItemid
		itemid_s = replace(itemid_s,",",chr(10))
		itemid_s = replace(itemid_s,chr(13),"")
		arrTemp = Split(itemid_s,chr(10))
		iA = 0
		do while iA <= ubound(arrTemp) 
			if trim(arrTemp(iA))<>"" then
				'상품코드 유효성 검사(2008.08.05;허진원)
				if isNumeric(trim(arrTemp(iA))) then
					arrItemid = arrItemid & trim(arrTemp(iA)) & ","
				end if
			end if
			iA = iA + 1
		loop

		itemid_s = left(arrItemid,len(arrItemid)-1)
	end if
	'// 재지정 끝

	'vWidth = CInt((100/vDepth))

	SET cDisp = New cDispCate
	cDisp.FCurrPage = vCurrpage
	cDisp.FPageSize = 2000
	cDisp.FRectDepth = vDepth
	cDisp.FRectCateCode = vCateCode
	cDisp.GetDispCateList()
%>

<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
function jsSaveDispCate(){
	if($('input[name="catename"]').val() == ""){
		alert("카테고리명을 입력하세요.");
		$('input[name="catename"]').focus();
		return;
	}
	if($('input[name="sortno"]').val() == ""){
		alert("정렬번호를 입력하세요.");
		$('input[name="sortno"]').focus();
		return;
	}
	if($.isNumeric($('input[name="sortno"]').val()) == false){
		alert("정렬번호는 숫자만 가능합니다.");
		$('input[name="sortno"]').val('');
		$('input[name="sortno"]').focus();
		return;
	}
	frmDispCate.submit();
}

function jsWriteCateCode(c,d,p){
	jsItemAllMoveFormClose();
	$.ajax({
			url: "display_cate_ajax.asp?catecode_s="+c+"&depth="+d+"&parentcatecode="+p+"",
			cache: false,
			success: function(message)
			{
				$("#catecodewritebox").empty().append(message);
				$("#catecodewritebox").show();
				jsFormCloseBtn(1);
			},
			error: function (xhr, status, error) {
				console.log(xhr.responseText);
			}
	});
}

function jsItemAllMove(){
	jsWriteFormClose();
	$.ajax({
			url: "display_cate_itemmove_ajax.asp",
			cache: false,
			success: function(message)
			{
				$("#itemallmovebox").empty().append(message);
				$("#itemallmovebox").show();
				jsFormCloseBtn(2);
			}
	});
}

function jsOldCateItemAllMove(){
	jsWriteFormClose();
	$.ajax({
			url: "display_oldcate_itemmove_ajax.asp",
			cache: false,
			success: function(message)
			{
				$("#itemallmovebox").empty().append(message);
				$("#itemallmovebox").show();
				jsFormCloseBtn(2);
			}
	});
}

function jsWriteFormClose(){
	$("#catecodewritebox").empty().append("");
	$("#catecodewritebox").hide();
	$("#closebtn1").hide();
}

function jsItemAllMoveFormClose(){
	$("#itemallmovebox").empty().append("");
	$("#itemallmovebox").hide();
	$("#closebtn2").hide();
}

function jsFormCloseBtn(g){
	if(g == 1){
		$("#closebtn1").show();
	}else{
		$("#closebtn2").show();
	}
}

function jsSelctCateShowItem(){
	dispcate_item.frmitem.IsThisCate.value = "o";
	dispcate_item.frmitem.cpg.value = 1;
	dispcate_item.frmitem.submit();
}

function jsCateCompleteDel(){
	var c = $('input[name="catecode"]').val();
	var cn = $('input[name="catename"]').val();
	if(confirm("code:"+c+", name:"+cn+" \n를 삭제하시겠습니까?\n\n※ 버튼 옆 주의사항을 꼭 확인하시기 바랍니다.") == true) {
		$('input[id="completedel"]').val("o");
		frmDispCate.submit();
	}
}

function jsSortCate(c,d){
	var sortcate = window.open('display_cate_sort.asp?depth='+d+'&catecode='+c+'','sortcate','width=400,height=500,scrollbars=yes, resizable=yes');
	sortcate.focus();
}

function jsEditLink(){
	location.href = "<%=CurrURLQ()%>#editlink";
}

function jsCatelistView(d,c){
	$('input[name="depth_s"]').val(d);
	$('input[name="catecode_s"]').val(c);

	catefrm.submit();
}

// 카테고리 다운로드
function jscategoryDown(){
	document.frmdown.target="cateproc";
	document.frmdown.action="/admin/CategoryMaster/displaycate/display_cate_exceldownload.asp";
	document.frmdown.submit();
}

</script>
<style type="text/css">
.box1 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FFF8F8; padding:7px 10px;}
.box2 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:5px; margin-top:5px;}
.box3 {width:<%=(286*vDepth)%>px; margin-top:5px;}
.box3 .subFirstBox {width:260px; border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; float:left; margin-left:0px;}
.box3 .subBox {width:260px; border:1px solid #CCCCCC; border-radius: 6px; padding:7px 7px; float:left; margin-left:5px;}
.box3 .subTTBox {border:0; border-radius: 6px; padding:3px 0; text-align:center; background-color:#888; color:#FFF; font-weight:bold;}
.box3 .subListBox {margin-top:5px;}
.box4 {border:1px solid #CCCCCC; border-radius: 6px; background-color:#FAFAFA; padding:7px 10px; ; margin-top:5px;}
.ttDep1 {background-color:#FAFAFA;}
.ttDep2 {background-color:#F5F5F5;}
.ttDep3 {background-color:#EFEFEF;}
.ttDep4 {background-color:#ECECEC;}
.ttDep5 {background-color:#E8E8E8;}
.ttDep6 {background-color:#E0E0E0;}
</style>
<div class="box1">
* 카테고리 수정하려면 [<span style="BACKGROUND-COLOR: #D4FFFF;">카테고리코드</span>] 를 클릭하세요.<br>
* 기본 정렬번호 : 카테고리 -> 99, 상품 -> 9999<br>
* <span style="border:1px solid #000000; BACKGROUND-COLOR: #FFF0F0; width:30px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;: 클릭한 카테고리,&nbsp;
<span style="border:1px solid #000000; BACKGROUND-COLOR: #CFCFCF; width:30px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;: 사용안함(전시안함) 카테고리<br>
* 각 2 Depth에는 맨 아래 부분에 BOOK 카테고리가 있어야 합니다.
</div>
<div class="box2">
	<table border="0" class="a">
	<tr id="lyrSbmBtn1">
		<td>
			<% If vDepth < 7 Then %><input type="button" value="<%=vDepth%> Depth 카테고리생성" onClick="jsWriteCateCode('','<%=vDepth%>','<%=vCateCode%>');"><% End If %>
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="카테고리내 전 상품 이동(New)" onClick="jsItemAllMove();">
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="[구→신] 카테고리 전환" onClick="jsOldCateItemAllMove();">
			&nbsp;&nbsp;&nbsp;
			<input type="button" value="선택한 카테고리 상품보기(하위뎁스포함)" onClick="jsSelctCateShowItem();">
			&nbsp;&nbsp;
			<input type="button" value="카테고리다운로드(엑셀)" onclick="jscategoryDown();">
		</td>
		<td align="right">
			<div id="closebtn1" style="display:none;"><input type="button" value="닫  기" onClick="jsWriteFormClose()"></div>
			<div id="closebtn2" style="display:none;"><input type="button" value="닫  기" onClick="jsItemAllMoveFormClose()"></div>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<form name="frmDispCate" method="post" action="display_cate_proc.asp" target="cateproc" style="margin:0px;">
			<input type="hidden" name="menupos" value="<%=request("menupos")%>">
			<div id="catecodewritebox" style="display:none;">
			</div>
			</form>
			<form name="frmItemAllMove" method="post" action="display_cate_itemmove_proc.asp" target="cateproc" style="margin:0px;">
			<input type="hidden" name="menupos" value="<%=request("menupos")%>">
			<div id="itemallmovebox" style="display:none;">
			</div>
			</form>
		</td>
	</tr>
	</table>
	<script>$("#lyrSbmBtn1 input").button();</script>
</div>
<div class="box3">
	<div class="subFirstBox ttDep1">
		<div class="subTTBox"><span style="padding-right:100px;"></span>1 Depth<span style="padding-left:57px;"><input type="button" value="정렬" style="height:16px;font-size:11px;" onClick="jsSortCate('',1);"></span></div>
		<div class="subListBox">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="#CCCCCC">
			<%
			Dim i, vTempDepth, vIsThisLine, vNowCateName, vBGcolor
			For i=0 To cDisp.FResultCount-1
				vIsThisLine = fnIsThisLine(cDisp.FItemList(i).FDepth,cDisp.FItemList(i).FCateCode,vCateCode)
				If vIsThisLine = "o" Then
					vNowCateName = vNowCateName & "[" & Right(cDisp.FItemList(i).FCateCode,3) & "]" & cDisp.FItemList(i).FCateName & " - "
				End IF
				
				If i=0 Then
					vTempDepth = cDisp.FItemList(i).FDepth
				End IF
	
				If vTempDepth <> cDisp.FItemList(i).FDepth Then
					Response.Write "	</table>" & vbCrLf &_
								"	</div>" & vbCrLf &_
								"</div>" & vbCrLf &_
								"<div class='subBox ttDep" & cDisp.FItemList(i).FDepth & "'>" & vbCrLf &_
								"	<div class='subTTBox'><span style='padding-right:100px;'></span>" & cDisp.FItemList(i).FDepth & " Depth<span style='padding-left:57px;'>" & _
								"<input type='button' value='정렬' style='height:16px;font-size:11px;' onClick='jsSortCate("&Left(cDisp.FItemList(i).FCateCode,3*(cDisp.FItemList(i).FDepth-1))&","&cDisp.FItemList(i).FDepth&");'></span></div>" & vbCrLf &_
								"	<div class='subListBox'>" & vbCrLf &_
								"	<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' class='a' bgcolor='#CCCCCC'>"
				End If
				
				vBGcolor = "#FFFFFF"
				If vIsThisLine = "o" Then
					vBGcolor = "#FFF0F0"
				End If
				If cDisp.FItemList(i).FUseYN = "N" Then
					vBGcolor = "#CFCFCF"
				End If
			%>
			<% If i = 0 Then %>
			<tr>
				<td bgcolor="#FFFFFF" width="260">
					<table width="100%" border=0 class=a>
					<tr>
						<td><a href="javascript:jsCatelistView('','');">Go 1 Depth</a></td>
						<td width="5" align="right"></td>
					</tr>
					</table>
				</td>
			</tr>
			<% End If %>
			<tr>
				<td bgcolor="<%=vBGcolor%>" width="260">
					<table width="100%" border=0 class=a>
					<tr>
						<td>
							<span onClick="jsWriteCateCode('<%=cDisp.FItemList(i).FCateCode%>','<%=cDisp.FItemList(i).FDepth%>','');" style="cursor:pointer;BACKGROUND-COLOR: #D4FFFF;">[<%=Right(cDisp.FItemList(i).FCateCode,3)%>]</span>
							<a href="javascript:jsCatelistView('<%=cDisp.FItemList(i).FDepth+1%>','<%=cDisp.FItemList(i).FCateCode%>');"><%=cDisp.FItemList(i).FCateName%></a>
						</td>
						<td width="5" align="right"><%=CHKIIF(vIsThisLine="o","▶","")%></td>
					</tr>
					</table>
				</td>
			</tr>
			<%
				vTempDepth = cDisp.FItemList(i).FDepth
			Next
			%>
			</table>
		</div>
	</div>
</div>
<form name="catefrm" method="get">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="depth_s" value="<%=vDepth%>">
<input type="hidden" name="catecode_s" value="<%=vCateCode%>">
<input type="hidden" name="notcatereg" value="<%=vNotCateReg%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="cdl" value="<%=cdl%>">
<input type="hidden" name="cdm" value="<%=cdm%>">
<input type="hidden" name="cds" value="<%=cds%>">
<input type="hidden" name="itemid_s" value="<%=itemid_s%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="keyword" value="<%=keyword%>">
<input type="hidden" name="sellyn" value="<%=sellyn%>">
<input type="hidden" name="usingyn" value="<%=usingyn%>">
<input type="hidden" name="danjongyn" value="<%=danjongyn%>">
<input type="hidden" name="limityn" value="<%=limityn%>">
<input type="hidden" name="sailyn" value="<%=sailyn%>">
<input type="hidden" name="deliverytype" value="<%=deliverytype%>">
<input type="hidden" name="sortDiv" value="<%=sortDiv%>">
<input type="hidden" name="pagesize" value="<%=pagesize%>">
</form>
<form name="frmdown" method="post" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<%
	Dim vParam
	vParam = "depth_s="&vDepth&"&catecode_s="&vCateCode&"&notcatereg="&vNotCateReg&"&makerid="&makerid&"&cdl="&cdl&"&cdm="&cdm&"&cds="&cds&"&itemid_s="&itemid_s&"&itemname="&itemname&"&keyword="&keyword&"&sellyn="&sellyn&"&usingyn="&usingyn&"&danjongyn="&danjongyn&"&limityn="&limityn&"&sailyn="&sailyn&"&deliverytype="&deliverytype&"&sortDiv="&sortDiv&"&pagesize="&pagesize&""
%>
<br style="clear:both;">
<div class="box4">
<a name="editlink" />
<input type="hidden" id="nowcatename" value="<% If vCateCode <> "" Then Response.Write Left(vNowCateName,(Len(vNowCateName)-3)) End If %>">
<iframe name="dispcate_item" id="dispcate_item" src="display_cate_item.asp?<%=vParam%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
</div>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="" id="cateproc" name="cateproc" width="100%" height="300" frameborder="0"></iframe>
<% else %>
	<iframe src="" id="cateproc" name="cateproc" width="0" height="0" frameborder="0"></iframe>
<% end if %>
<% SET cDisp = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->