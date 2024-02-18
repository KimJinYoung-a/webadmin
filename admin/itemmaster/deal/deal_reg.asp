<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/itemmaster/deal/deal_reg.asp
' Description :  딜 이벤트 등록
' History : 2017.08.23 정태훈
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/newdealManageCls.asp"-->
<%
CONST CBASIC_IMG_MAXSIZE = 560   'KB
Dim idx
idx = requestCheckVar(Request("idx"),10)
If idx="" Then
	Dim oDealMax
	set oDealMax = New ClsDeal
	oDealMax.fnGetMAXDealMasterNum
	idx=oDealMax.FMasterIDX
	Set oDealMax=Nothing
Response.redirect "/admin/itemmaster/deal/deal_reg.asp?idx="&idx
Response.End
End If
%>
<script type="text/javascript">
<!--
	function TnViewDivSelect(viewdiv){
		if(viewdiv==1){
			$("#datearea").css("display","none");
		}else{
			$("#datearea").css("display","");
		}
	}

	function TnSearchObjOpenWin(){
		var winpop = window.open('/admin/itemmaster/deal/pop_deal_addItems.asp?idx=<%=idx%>&stype=w','winpop','width=1450,height=800,scrollbars=yes,resizable=yes');
	}

	function SubmitSave(frm){
		if(frm.itemname.value=="")
		{
			alert("상품명을 입력해주세요.");
			frm.itemname.focus();
			return false;
		}
		else if(frm.itemname.value.length>50)
		{
			alert("50자 이내로 상품명을 입력해주세요.");
			frm.itemname.focus();
			return false;
		}
		else if(!frm.viewdiv[0].checked && !frm.viewdiv[1].checked)
		{
			alert("노출 기간을 선택해주세요.");
			return false;
		}
		else if(frm.viewdiv[1].checked && (frm.startdate.value=="" || frm.enddate.value==""))
		{
			alert("노출 기간을 설정해주세요.");
			return false;
		}
		else if(!frm.isusing[0].checked && !frm.isusing[1].checked)
		{
			alert("사용 여부를 선택해주세요.");
			return false;
		}
		else if(frm.itemid.value=="")
		{
			alert("대표상품을 선택해주세요.");
			frm.itemid.focus();
			return false;
		}
		else if(frm.mastersellcash.value=="")
		{
			alert("대표 가격을 입력해주세요.");
			frm.mastersellcash.focus();
			return false;
		}
		else if(frm.masterdiscountrate.value=="")
		{
			alert("대표 할인율을 입력해주세요.");
			frm.masterdiscountrate.focus();
			return false;
		}
		else if($("#tbl_DispCate tr").length<1)
		{
			alert("전시 카테고리를 추가해주세요.");
			frm.catecode.focus();
			return false;
		}
		else if(frm.keywords.value=="")
		{
			alert("검색 키워드를 입력해 주세요.");
			frm.keywords.focus();
			return false;
		}
		else
		{
			if(confirm("입력하신 정보로 딜상품을 등록하시겠습니까?"))
			{
				frm.target="FrameCKP";
				frm.action="dodealinfo_process.asp";
				frm.submit();
			}
		}
	}

    function TnDealSaveAPICall(itemid){
		document.frm.target="FrameCKP";
		document.frm.tempitemid.value=itemid;
		document.frm.action="<%= ItemUploadUrl %>/linkweb/items/deal_itemregisterTempWithImage_process.asp";
		frm.submit();
    }


	function TnMasterItemSelect(itemid){
		if(document.frm.itemname.value=="")
		{
			document.frm.itemname.value=$("#itemcode option:selected").text();
		}
		$.ajax({
			url: "selectdealitemkeywords.asp?itemid="+itemid,
			cache: false,
			async: false,
			success: function(message) {
				//alert(message);
				if(message!="") {
					$('#keywords').val(message);
				} else {
					alert("제공 할 정보가 없습니다.");
				}
			}
		});
	}

	// 기본정보 수정
	function editItemBasicInfo(itemid) {
		var param = "itemid=" + itemid + "&menupos=<%= menupos %>";
		popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	// 기본정보 수정
	function fnSaleInfo() {
		popwin = window.open('/admin/shopmaster/sale/saleList.asp?menupos=290' ,'saleinfo','width=1100,height=600,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	function onlyNumerSet(text){
		if(window.event.keyCode < 48 || window.event.keyCode > 57) {
			return false;
		}
	}

	function fnPaste() {
		var regex = /\D/ig;
		if (regex.test(window.clipboardData.getData("text"))) {
			return false;
		} else {
			return true;
		}
	}

	function jsSetImg(sName, sSpan){ 
		var winImg;
		winImg = window.open('pop_deal_uploadimg.asp?yr=<%=Year(now())%>&sName='+sName+'&sSpan='+sSpan+'&wid=800&hei=1600','popImg','width=370,height=150');
		winImg.focus();
	}

//-->
</script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" method="post" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="auser" value="<%=session("ssBctId")%>">
<input type="hidden" name="catecnt" id="catecnt">
<input type="hidden" name="mode" value="reg">
<input type="hidden" name="tempitemid">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>딜 기본 정보</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">상품명<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="itemname" id="itemname" size="80" maxlength="120" value="" class="text">
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">노출 기간<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			 <input type="radio" name="viewdiv" id="viewdiv" value="1" onClick="TnViewDivSelect(1)" checked>상시딜 <input type="radio" name="viewdiv" id="viewdiv" value="2" onClick="TnViewDivSelect(2)">기간딜
			 <span id="datearea" style="display:none">
				<input id="startdate" name="startdate" value="<%=Date()%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="shour" size="2" class="text" value="00" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">:<input type="text" name="sminute" size="2" class="text" value="00" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');"> ~
				<input id="enddate" name="enddate" value="<%=DateAdd("D",14,Date())%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iED_trigger" border="0" style="cursor:pointer" align="absmiddle" /> <input type="text" name="ehour" size="2" class="text" value="23" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">:<input type="text" name="eminute" size="2" class="text" value="59" maxlength="2" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "startdate", trigger    : "iSD_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "enddate", trigger    : "iED_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			 </span>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사용여부<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="radio" name="isusing" id="isusing" value="Y" checked>사용 <input type="radio" name="isusing" id="isusing" value="N">사용 안함
		</td>
	</tr>
</table>
<p>&nbsp;</p>
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td align="left" colspan="2" bgcolor="<%= adminColor("tabletop") %>"><B>노출 상품 정보</B></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100">상품목록<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="button" class="button" style="width:105;" value="검색" onclick="TnSearchObjOpenWin('Just1Day_list.asp');">&nbsp;<b style="color:red">*</b>딜 상품을 검색하여 추가해 주세요.
			<div id="divForm">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
			<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
				<td>순서</td>
				<td>상품코드</td>
				<td>상품명</td>
				<td>판매가</td>
				<td>매입가</td>
				<td>할인율</td>
			</tr>
			</table>
			</div>
			<div id="divFrm3"></div>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 상품<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<select name="itemid"  id="itemcode" disabled  onChange="TnMasterItemSelect(this.value);">
				<option value="" selected>상품을 추가해 주세요.</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 가격, 할인<br>아이템 코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="saleitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
			<input type="text" name="discountitemid" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">
			(수기로 아이템 코드를 입력해도 가격,할인정보를 수정 할 수 있습니다.)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 가격<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="mastersellcash" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">&nbsp;<input type="button" value="가져오기" class="button" onClick="fnGetMinPricevalue()" id="saleper1" name="saleper1" style="display:none"><!-- &nbsp;<input type="checkbox" name="pricesdash" value="Y">"~"노출 선택 (예: 19,900원~) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표 할인율<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="masterdiscountrate" id="masterdiscountrate" value="" size="10" class="text" maxlength="10" pattern="[A-Za-z0-9]*" onKeyPress="return onlyNumerSet(this);" onPaste="return fnPaste();" onkeyup="this.value=this.value.replace(/[\ㄱ-ㅎㅏ-ㅣ가-힣]/g, '');">&nbsp;<input type="button" value="가져오기" class="button" onClick="fnGetMaxSalevalue()" id="saleper2" name="saleper2" style="display:none"><!-- &nbsp;<input type="checkbox" name="sailsdash" value="Y">"~"노출 선택 (예: ~77%) -->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">전시 카테고리<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<table class="a">
			<tr>
				<td id="lyrDispList"><table class="a" id="tbl_DispCate"></table></td>
				<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
			</tr>
			</table>
			<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
		</td>
	</tr>
	<tr align="left">
	<td height="30" width="15%" bgcolor="<%= adminColor("tabletop") %>">구매 가능 연령 </td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="adultType" value="0" checked>전체연령</label>
		<label><input type="radio" name="adultType" value="1" >구매시성인인증</label>
		<label><input type="radio" name="adultType" value="2" >미성년 조회 불가</label>
	</td>
</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">검색 키워드<b style="color:red">*</b></td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="keywords" id="keywords" size="80" maxlength="250" value="" class="text"> (콤마로구분 ex: 커플,티셔츠,조명)
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">요약 이미지</td>
		<td bgcolor="#FFFFFF">
			<input class="button" type="button" value="이미지 불러오기" onClick="jsSetImg('dealcontents','spandealcontents');"/>
			(선택,800X1600, Max 800KB,jpg,gif)
			<div id="spandealcontents"></div>
			<input type="hidden" name="addimggubun" value="1">
			<input type="hidden" name="addimgdel" value="">
			<input type="hidden" name="dealcontents" value="">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">비고</td>
		<td bgcolor="#FFFFFF">
			<textarea name="work_notice" rows="18" class="textarea" style="width:99%" id="[on,off,off,off][상품설명]"></textarea>
		</td>
	</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" height="25">
        <td valign="bottom" align="center">
			<input type="button" value="등록" class="button" onClick="SubmitSave(this.form)">
			<input type="button" value="취소" class="button" onClick="fnCancel()">
        </td>
    </tr>
</table>
</form>
<% if (application("Svr_Info")	= "Dev") then %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% else %>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="300" height="100"></iframe>
<% end if %>
<script type="text/javascript">
<!--
	function fnCancel(){
		if(confirm("입력하신 정보를 저장하지 않고 취소하시겠습니까?")){
			location.href="/admin/itemmaster/deal/index.asp";
		}
	}
	//상품 최대 할인율 접수
	function fnGetMaxSalevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMaxItemSalePer.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.masterdiscountrate.value=splitmessage[0];
					document.frm.discountitemid.value=splitmessage[1];
				} else {
					alert("상품이 없거나 할인중인 상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	//상품 최저가 접수
	function fnGetMinPricevalue() {
		var idx = document.frm.idx.value;
		$.ajax({
			type: "POST",
			url: "ajaxGetDealMinItemPrice.asp",
			data: "idx="+idx,
			cache: false,
			success: function(message) {
				var splitmessage = message.split("|")
				if(message) {
					document.frm.mastersellcash.value=splitmessage[0];
					document.frm.saleitemid.value=splitmessage[1];
				} else {
					alert("상품이 없습니다.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 레이어에서 전시카테고리 추가
	function addDispCateItem(dcd,cnm,div,dpt) {
		// 기존에 값에 중복 카테고리 여부 검사
		if(tbl_DispCate.rows.length>=2)	{
			alert("전시 카테고리는 최대 2개까지 입력가능합니다.");
			return false;
		}
		else
		{
			if(tbl_DispCate.rows.length>0)	{
				if(tbl_DispCate.rows.length>1)	{
					for(l=0;l<document.all.isDefault.length;l++)	{
						if((document.all.catecode[l].value==dcd)) {
							alert("이미 지정된 같은 카테고리가 있습니다..");
							return;
						}
					}
				}
				else {
					if((document.all.catecode.value==dcd)) {
						alert("이미 지정된 같은 카테고리가 있습니다..");
						return;
					}
				}
			}
		}
		
		// 행추가
		var oRow = tbl_DispCate.insertRow();
		oRow.onmouseover=function(){tbl_DispCate.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="y") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='isDefault' value='y'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='isDefault' value='n'>";
		}
		$(cnm).each(function(i){
			if(dpt>i) {
				if(i>0) oCell2.innerHTML += " >> ";
				oCell2.innerHTML += $(this).text();
			}
		});
		oCell2.innerHTML += "<input type='hidden' name='catecode' value='" + dcd + "'>";
		oCell2.innerHTML += "<input type='hidden' name='catedepth' value='" + dpt + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delDispCateItem()' align=absmiddle>";
		$("#lyrDispCateAdd").fadeOut();
		$("#catecnt").val($("#catecnt").val()+1);
		//상품속성 출력
		printItemAttribute();
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);
			$("#catecnt").val($("#catecnt").val()-1);
			//상품속성 출력
			printItemAttribute();
		}
	}

	function printItemAttribute() {
		var arrDispCd="";
		$("input[name='catecode']").each(function(i){
			if(i>0) arrDispCd += ",";
			arrDispCd += $(this).val();
		});
		$.ajax({
			url: "/common/module/act_ItemAttribSelect.asp?itemid=0&arrDispCate="+arrDispCd,
			cache: false,
			success: function(message) {
				$("#lyrItemAttribAdd").empty().append(message);
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	function CheckImage(img, filesize, imagewidth, imageheight, extname, fsize)
	{
		var ext;
		var filename;

		filename = img.value;
		if (img.value == "") { return false; }

		if (CheckExtension(filename, extname) != true) {
			alert("이미지화일은 다음의 화일만 사용하세요.[" + extname + "]");
			ClearImage(img,fsize,imagewidth,imageheight);
			return false;
		}

		return true;
	}

	function ClearImage2(img,fsize,wd,ht) {
		img.outerHTML="<input type='file' name='" + img.name + "' onchange=\"CheckImage(this, <%= CBASIC_IMG_MAXSIZE %>, "+wd+", "+ht+", 'jpg,gif', "+ fsize +");\" class='text' size='"+ fsize +"'>";
	}
//-->
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->