<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품수정
' History : 서동석 생성
'			2016.07.06 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, oitem, oitemvideo
dim makerid, rentalItemFlag

itemid = requestCheckvar(request("itemid"),10)
makerid = requestCheckvar(request("makerid"),32)
menupos = requestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

''연관상품 목록 접수
dim strItemRelation
strItemRelation = GetItemRelationStr(itemid)

'세일마진
dim sailmargine, orgmargine, margine

''수정
if oitem.FOneItem.Fsailprice<>0 then
	sailmargine = 100-CLng(oitem.FOneItem.Fsailsuplycash/oitem.FOneItem.Fsailprice*100*100)/100
else
	sailmargine = 0
end if

if oitem.FOneItem.Forgprice<>0 then
	orgmargine = 100-CLng(oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice*100*100)/100
else
	orgmargine = 0
end if

if oitem.FOneItem.Fsellcash<>0 then
	margine = 100-CLng(oitem.FOneItem.Fbuycash/oitem.FOneItem.Fsellcash*100*100)/100
else
	margine = 0
end if

'// 렌탈 상품은 일단 테스트로 위탁 유저만 노출함
If C_ADMIN_AUTH Then
	rentalItemFlag = true
Else
	rentalItemFlag = true
End If
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript' src="/js/ckeditor/ckeditor.js"></script>
<script type="text/javascript">
$(function(){
	// 로딩후 상품속성 내용 출력
	printItemAttribute();
});

function printItemAttribute() {
	var arrDispCd="";
	$("input[name='catecode']").each(function(i){
		if(i>0) arrDispCd += ",";
		arrDispCd += $(this).val();
	});
	$.ajax({
		url: "/common/module/act_ItemAttribSelect.asp?itemid=<%=itemid%>&arrDispCate="+arrDispCd,
		cache: false,
		success: function(message) {
			$("#lyrItemAttribAdd").empty().append(message);
		}
		,error: function(err) {
			alert(err.responseText);
		}
	});
}

function UseTemplate() {
	window.open("/common/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=600, height=500, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 카테고리등록
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.itemreg;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}


// ============================================================================
// 저장하기
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

	if (!$("input[name='isDefault'][value='y']").length&&$("input[name='isDefault']").length){
		alert("[기본] 전시 카테고리를 선택하세요.\n※ [추가] 전시 카테고리만 넣을 수 없습니다.");
		return;
	}

	// 카테고리 지정여부 검사
	if(tbl_Category.rows.length>0)	{
		if(tbl_Category.rows.length>1)	{
			var chk=0;
			for(l=0;l<document.all.cate_div.length;l++)	{
				if(document.all.cate_div[l].value=="D") chk++;
			}
			if(chk==0) {
				alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
				return;
			} else if(chk>1) {
				alert("카테고리에 기본 카테고리를 한개만 선택해주세요.");
				return;
			}
		}
		else {
			if(document.all.cate_div.length){
				if(document.all.cate_div[0].value!="D") {
					alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
					return;
				}
			} else {
				if(document.all.cate_div.value!="D") {
					alert("카테고리에 기본 카테고리를 선택해주세요.\n※기본 카테고리는 필수항목입니다.");
					return;
				}
			}
		}
	} else {
		alert("카테고리를 선택해주세요.");
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }
    
    //업체배송만 주문제작 가능.
    <% if oitem.FOneItem.Fmwdiv <> "U" then %>
    if (itemreg.itemdiv[1].checked){
        alert('주문 제작상품은 업체배송인경우만 가능합니다.');
        itemreg.itemdiv[0].focus();
        return;
    }
    <%else%>//매입,위탁만 단독(예약) 주문설정 가능
    	if(itemreg.reserveItemTp[1].checked){
    		if(!confirm("단독(예약)구매상품은 다른 상품과 같이 구매가 불가합니다.\n단독 구매상품으로 변경하시겠습니까?")){
				itemreg.reserveItemTp[0].focus();
				return;
			};
    	}
    <% end if %>
    
    //상품명 길이체크 추가 64Byte
	if (getByteLength(itemreg.itemname.value)>64){
	    alert("상품명은 최대 64byte 이하로 입력해주세요.(한글32자 또는 영문64자)");
		itemreg.itemname.focus();
		return;
	}

	//상품 품목정보
    if (!itemreg.infoDiv.value){
        alert('상품에 해당하는 품목을 선택해주십시요.');
        itemreg.infoDiv.focus();
        return;
    } else if(itemreg.infoDiv.value=="35") {
    	if(!itemreg.itemsource.value) {
	        alert('상품의 재질을 입력해주세요.');
	        itemreg.itemsource.focus();
	        return;
    	}
    	if(!itemreg.itemsize.value) {
	        alert('상품의 크기를 입력해주세요.');
	        itemreg.itemsize.focus();
	        return;
    	}
    }

	//안전인증정보.
    if (itemreg.safetyYn[0].checked){
  		if($("#real_safetynum").val() == ""){
  			alert("안전인증구분을 선택하고 인증번호를 입력후 추가버튼을 클릭해주세요.");
  			return;
  		}
    }

    if(confirm("상품을 올리시겠습니까?") == true){
		<% ''안전인증 api로 조회 후 받은 데이터 db저장 후 생성idx값 받아 셋팅 %>
		if(itemreg.safetyYn[0].checked) {
			$("#real_safetyidx").val(jsCallAPIsafety($("#real_safetynum").val(),"o",$("#real_safetydiv").val()));
		}

        itemreg.submit();
    }

}

function getByteLength(inputValue) {
     var byteLength = 0;
     for (var inx = 0; inx < inputValue.length; inx++) {
         var oneChar = escape(inputValue.charAt(inx));
         if ( oneChar.length == 1 ) {
             byteLength ++;
         } else if (oneChar.indexOf("%u") != -1) {
             byteLength += 2;
         } else if (oneChar.indexOf("%") != -1) {
             byteLength += oneChar.length/3;
         }
     }
     return byteLength;
}

// ============================================================================
	// 카태고리 선택 팝업
	function popCateSelect(iid){
		var popwin = window.open("/common/module/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var designerid = document.all.itemreg.makerid.value;
		if(designerid == ""){
			alert("업체를 선택하세요.");
			return;
		}
		
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/common/module/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
			cache: false,
			success: function(message) {
				$("#lyrDispCateAdd").empty().append(message).fadeIn();
			}
			,error: function(err) {
				alert(err.responseText);
			}
		});
	}

	// 팝업에서 선택 카테고리 추가
	function addCateItem(lcd,lnm,mcd,mnm,scd,snm,div)
	{
		// 기존에 값에 중복 카테고리 여부 검사 - 플라워 전국배송은 제외함;
		if(tbl_Category.rows.length>0)	{
			if(tbl_Category.rows.length>1)	{
				for(l=0;l<document.all.cate_div.length;l++)	{
				    if (!((document.all.cate_large[l].value=="110")&&(document.all.cate_mid[l].value=="060"))){
    					if((document.all.cate_large[l].value==lcd)&&(document.all.cate_mid[l].value==mcd)) {
    						alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n기존 카테고리를 삭제하고 다시 선택해주세요.");
    						return;
    					}
    				}
				}
			}
			else {
			    if (!((document.all.cate_large.value=="110")&&(document.all.cate_mid.value=="060"))){
    				if((document.all.cate_large.value==lcd)&&(document.all.cate_mid.value==mcd)) {
    					alert("같은 중분류에 이미 지정된 카테고리가 있습니다.\n※기존 카테고리를 삭제하고 다시 선택해주세요.");
    					return;
    				}
    			}
			}
		}
		
		// 행추가
		var oRow = tbl_Category.insertRow();
		oRow.onmouseover=function(){tbl_Category.clickedRowIndex=this.rowIndex};

		// 셀추가 (구분,카테고리,삭제버튼)
		var oCell1 = oRow.insertCell();		
		var oCell2 = oRow.insertCell();
		var oCell3 = oRow.insertCell();

		if(div=="D") {
			oCell1.innerHTML = "<font color='darkred'><b>[기본]<b></font><input type='hidden' name='cate_div' value='D'>";
		} else {
			oCell1.innerHTML = "<font color='darkblue'>[추가]</font><input type='hidden' name='cate_div' value='A'>";
		}
		oCell2.innerHTML = lnm + " >> " + mnm + " >> " + snm
					+ "<input type='hidden' name='cate_large' value='" + lcd + "'>"
					+ "<input type='hidden' name='cate_mid' value='" + mcd + "'>"
					+ "<input type='hidden' name='cate_small' value='" + scd + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delCateItem()' align=absmiddle>";
	}

	// 레이어에서 전시카테고리 추가
	function addDispCateItem(dcd,cnm,div,dpt) {
		// 기존에 값에 중복 카테고리 여부 검사
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

		//상품속성 출력
		printItemAttribute();
	}

	// 선택 카테고리 삭제
	function delCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}

	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			printItemAttribute();
		}
	}

function checkItemDiv(comp){
    var frm = comp.form;
    
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
        }
    }
    
    //주문제작 상품인경우.
    if (frm.itemdiv[1].checked){
        if (frm.reqMsg.checked){
            frm.itemdiv[1].value="06";
        }else{
            frm.itemdiv[1].value="16";
        }
    }

    //렌탈 상품인경우.
    if (frm.itemdiv[7].checked){
		frm.reserveItemTp[1].checked = true;
    }	
}

//품목 선택 / 품목내용 표시
function chgInfoDiv(v) {
	$("#itemInfoList").empty();

	if(v=="") {
		$("#itemInfoCont").hide();
	} else {
		$("#itemInfoCont").show();

		var str = $.ajax({
			type: "POST",
			url: "act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v,
			dataType: "html",
			async: false
		}).responseText;
	
		if(str!="") {
			$("#itemInfoList").html(str);
		}
	}

	if(v=="35") {
		$("#lyItemSrc").show();
		$("#lyItemSize").show();
	} else {
		$("#lyItemSrc").hide();
		$("#lyItemSize").hide();
	}

	// 안전인증체크. 전안법
	jsSafetyCheck('','');
}

//단순 라디오 선택자
function chgInfoChk(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
}

//문구 라디오 선택자
function chgInfoSel(fm) {
	$(fm).parent().parent().find('[name="infoChk"]').val($(fm).val());
	$(fm).parent().parent().find('[name="infoCont"]').val($(fm).attr("msg"));

	if($(fm).val()=="Y") {
		$(fm).parent().parent().find('[name="infoCont"]').removeAttr("readonly");
		$(fm).parent().parent().find('[name="infoCont"]').removeClass("text_ro");
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text");
	} else {
		$(fm).parent().parent().find('[name="infoCont"]').attr("readonly", true);
		$(fm).parent().parent().find('[name="infoCont"]').addClass("text_ro");
	}
}

//상품군에 따른 원산지 설명 표기
function jsSetArea(iValue){ 
	var i;
	for(i=0;i<=4;i++) { 
 		eval("document.all.dvArea"+i).style.display = "none";
	}
 	eval("document.all.dvArea"+iValue).style.display = "";
} 

function jsCallAPIsafety(certnum,isSave,safetydiv){
	var returnmsg = "";
	$.ajax({
		url: "/admin/itemmaster/safety_api_auth_proc.asp?itemid=<%=itemid%>&issave="+isSave+"&certnum="+certnum+"&safetydiv="+safetydiv+"&statusmode=real",
		cache: false,
		async: false,
		success: function(message)
		{
			returnmsg = message;
		}
	});
	return returnmsg;
}

//전시카테고리(안전인증값)에 따른 alert 메세지.
function jsAlertCatecodeSafety(){
	var auth_go_catecode = "";
	if(typeof itemreg.catecode != "undefined"){
		if(itemreg.catecode.length == undefined){
			auth_go_catecode = itemreg.catecode.value;
		}else{
			for(si=0; si<itemreg.catecode.length; si++){
				auth_go_catecode = auth_go_catecode + itemreg.catecode[si].value + ",";
			}
		}
		
		if(auth_go_catecode != ""){
			$("#auth_go_catecode").val(auth_go_catecode);
			
			var ccode = $("#auth_go_catecode").val();
			$.ajax({
					url: "/common/item/catecode_safety_info_ajax.asp?catecode="+ccode,
					cache: false,
					async: false,
					success: function(msgc)
					{
						if(msgc != ""){
							msgc = msgc.replace(/br/gi,"\n");
							alert(msgc);
						}
					}
			});
		}
	}else{
		alert("전시카테고리를 선택해주세요.");
	}
}

//추가된 안전인증 리스트 개별 삭제
function jsSafetyDivListDel(listnum){
	var realvalue = $("#real_safetydiv").val();
	var jbSplit = $("#real_safetydiv").val().split(",");
	var jbSplitnum = $("#real_safetynum").val().split(",");
	var resultDiv = "";
	var resultNum = "";
	var del_safetynum = "";
	var del_safetydiv = "";
	
	for(var i in jbSplit){
		if(jbSplit[i] != listnum){
			resultDiv = resultDiv + jbSplit[i] + ",";
			resultNum = resultNum + jbSplitnum[i] + ",";
		}else{
			del_safetynum = jbSplitnum[i];
			del_safetydiv = jbSplit[i];
		}
	}
	
	if(resultDiv.substr(resultDiv.length-1, 1) == ","){
		resultDiv = resultDiv.substr(0, resultDiv.length-1);
		resultNum = resultNum.substr(0, resultNum.length-1);
	}
	$("#real_safetydiv").val(resultDiv);
	$("#real_safetynum").val(resultNum);
	
	$("#l"+listnum+"").remove();
	
	var tmp_num = $("#real_safetynum_delete").val();
	var tmp_div = $("#real_safetydiv_delete").val();
	if(tmp_num == ""){
		$("#real_safetynum_delete").val(del_safetynum);
		$("#real_safetydiv_delete").val(del_safetydiv);
	}else{
		$("#real_safetynum_delete").val(tmp_num + "," + del_safetynum);
		$("#real_safetydiv_delete").val(tmp_div + "," + del_safetydiv);
	}
}

// 브랜드ID 변경
function fnChangeBrandID() {
//
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
<tr height="10" valign="bottom">
	<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
	<font color="red"><strong>상품 기본정보 수정</strong></font></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<br><b>등록된 상품의 기본정보를 수정합니다.</b>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr  height="10"valign="top">
	<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>

<p>

<form name="itemreg" method="post" action="/admin/itemmaster/itemmodify_Process.asp" onsubmit="return false;" style="margin:0;">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="ItemBasicInfo">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td background="/images/tbl_blue_round_02.gif" colspan="2"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="left">기본정보</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		<%= oitem.FOneItem.Fitemid %>
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="미리보기" class="button" onclick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF">등록/판매일 :</td>
	<td width="35%" bgcolor="#FFFFFF">
		제품등록일 : <%= oitem.FOneItem.FRegDate %>
		<%
			if oitem.FOneItem.FsellSTDate<>"" then
				Response.Write "<br />판매시작일 : " & oitem.FOneItem.FsellSTDate
			elseif oitem.FOneItem.Fsellreservedate<>"" then
				Response.Write "<br />판매예정일 : " & oitem.FOneItem.Fsellreservedate
			end if
		%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="관리 브랜드ID">브랜드ID :</td> 
	<td width="35%" bgcolor="#FFFFFF">
		<% 'NewDrawSelectBoxDesignerChangeMargin "makerid", oitem.FOneItem.Fmakerid, "marginData", "fnChangeBrandID" %>
		<% drawSelectBoxDesignerWithName "makerid", oitem.FOneItem.Fmakerid %>
	</td>
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에서 표시될 브랜드(없으면 관리 브랜드 적용됨)">표시 브랜드 :</td>
	<td width="35%" bgcolor="#FFFFFF">
	<%
		drawSelectBoxDesignerWithName "frontMakerid", oitem.FOneItem.FfrontMakerid

		'표시브랜드 삭제 버튼
		response.Write "&nbsp;<input type=""button"" class=""button"" value=""제거"" onClick=""this.form.frontMakerid.value='';"">"
	%>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드명 :</td>
	<td bgcolor="#FFFFFF" colspan="3" id="txtBrandName"><%=oitem.FOneItem.Fbrandname%></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemname" maxlength="64" size="60" class="text" id="[on,off,off,off][상품명]" value="<%= Replace(oitem.FOneItem.Fitemname,"""","&quot;") %>">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">영문상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][영문상품명]" value="<%= Replace(oitem.FOneItem.FitemnameEng,"""","&quot;") %>">&nbsp;
		<input type="button" value="다국어 정보 <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"등록","수정")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="재고/매출 등의 관리 카테고리" style="cursor:help;">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="추가" class="button" onClick="popCateSelect('<%=oitem.FOneItem.Fitemid%>')"></td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">일반상품</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">주문 제작상품</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문시 이니셜등 제작문구가 필요한경우 체크)</font>
		<br>
		<label><input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">티켓상품</label>
		<label><input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present상품</label>
		<label><input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">여행상품</label>

		<!--% if oitem.FOneItem.Fitemdiv ="07" then %--> <!-- 2014년이전 단독구매 상품 > reserveItemTp=1 / 현재는 구매제한(회원당 구매 제한) -->
			<label><input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">구매제한상품</label>
		<!--% end if %-->
		<% if oitem.FOneItem.Fitemdiv ="82" then %>
			<label><input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">마일리지샵 상품</label>
		<% end if %>

		<label><input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">정기구독상품</label>

		<% If rentalItemFlag Then %>
			<label><input type="radio" name="itemdiv" value="30" <%=chkIIF(oitem.FOneItem.Fitemdiv="30","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">렌탈상품<font color=red>(렌탈상품은 반드시 단독(예약)구매상품으로 등록하셔야 합니다.)</font></label>
		<% End If %>
		<label><input type="radio" name="itemdiv" value="23" <%=chkIIF(oitem.FOneItem.Fitemdiv="23","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">B2B상품</label>
		<label><input type="radio" name="itemdiv" value="17" <%=chkIIF(oitem.FOneItem.Fitemdiv="17","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">마케팅전용상품</label>
		<label><input type="radio" name="itemdiv" value="11" <%=chkIIF(oitem.FOneItem.Fitemdiv="11","checked","")%> onClick="document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">상품권상품</label>
	</td>
	<td  bgcolor="#FFFFFF">
	    <div id="lyRequre" style="<%=chkIIF((oitem.FOneItem.Fitemdiv ="06") or (oitem.FOneItem.Fitemdiv ="16"),"","display:none;")%>padding-left:22px;">
		예상제작소요일 <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][예상제작소요일]">일
		<font color="red">(상품발송전 상품제작 기간)</font>
		</div>
	</td>
</tr>
<!-- 개발중단 2017.10.17 정윤정(염기호 기획) -->
<!--<tr>
	<td height="30" width="15%" bgcolor="#DDDDFF"> 브랜드 선물포장 : </td>
	<td bgcolor="#FFFFFF" colspan="3">	 
	<span style="margin-right:10px;"><input type="checkbox" name="chkMsg" value="Y" <%if oitem.FOneItem.FaddMsg ="Y" then%>checked<%end if%>> 메시지 첨부</span>
	<span  style="margin-right:10px;"><input type="checkbox" name="chkCarve"  value="Y" <%if oitem.FOneItem.FaddCarve ="Y" then%>checked<%end if%>> 각인 서비스</span>
	<span  style="margin-right:10px;"><input type="checkbox"  name="chkBox"  value="Y" <%if oitem.FOneItem.FaddBox ="Y" then%>checked<%end if%>>  박스포장</span>
	<span style="margin-right:10px;"><input type="checkbox"  name="chkSet"  value="Y"<%if oitem.FOneItem.FaddSet ="Y" then%>checked<%end if%>>  선물세트</span>
	<span  style="margin-right:10px;"><input type="checkbox"  name="chkCustom"  value="Y" <%if oitem.FOneItem.FaddCustom ="Y" then%>checked<%end if%>>  주문제작 </span>
	</td>
</tr>-->
<!---// ---------------------->
<!--% if (oitem.FOneItem.IsReserveOnlyItem) then %-->
<!-- 설정은 시스템팀 only 2012/03/26 추가-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">단독(예약)구매 :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
	    <%if isNull(oitem.FOneItem.FreserveItemTp) then oitem.FOneItem.FreserveItemTp=0 %>
	    <label><input type="radio" name="reserveItemTp" value="0" <%=chkIIF(oitem.FOneItem.FreserveItemTp="0" And oitem.FOneItem.Fitemdiv<>"30","checked","")%>>일반</label>
		<label><input type="radio" name="reserveItemTp" value="1" <%=chkIIF(oitem.FOneItem.FreserveItemTp="1" or oitem.FOneItem.Fitemdiv="30","checked","")%>>단독(예약)구매상품</label>
	</td>
</tr>
<!--% end if %-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">텐바이텐 독점 :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
		<label><input type="radio" name="tenOnlyYn" value="Y" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="Y","checked","")%>>독점상품</label>
		<label><input type="radio" name="tenOnlyYn" value="N" <%=chkIIF(oitem.FOneItem.FtenOnlyYn="N","checked","")%>>일반상품</label>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">구매 가능 연령 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<label><input type="radio" name="adultType" value="0" <%=chkIIF(oitem.FOneItem.FadultType=0,"checked","")%>>전체연령</label>
		<label><input type="radio" name="adultType" value="1" <%=chkIIF(oitem.FOneItem.FadultType=1,"checked","")%>>미성년 조회불가</label>
		<label><input type="radio" name="adultType" value="2" <%=chkIIF(oitem.FOneItem.FadultType=2,"checked","")%>>구매시 성인인증</label>
	</td>
</tr>
<!--<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">선착순 결제 상품 :</td>
	<td width="85%" bgcolor="#FFFFFF" colspan="3"> 
		<label><input type="radio" name="availPayType" value="9" <%=chkIIF(oitem.FOneItem.FavailPayType="9","checked","")%>>선착순</label>
		<label><input type="radio" name="availPayType" value="8" <%=chkIIF(oitem.FOneItem.FavailPayType="8","checked","")%>>저스트원데이</label>
		<label><input type="radio" name="availPayType" value="0" <%=chkIIF(oitem.FOneItem.FavailPayType="0","checked","")%>>일반</label> 
	</td>
</tr>-->
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품카피 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="designercomment" size="60" maxlength="128" class="text" id="[off,off,off,off][상품카피]" value="<%= Replace(oitem.FOneItem.Fdesignercomment,"""","&quot;") %>">
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" class="text" id="[on,off,off,off][상품무게]" style="text-align:right" value="<%= oitem.FOneItem.FitemWeight %>">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
	<td bgcolor="#FFFFFF" colspan="3"> 
		 <p> 
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="0" <%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 식품 외</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="1" <%if oitem.FOneItem.Fsourcekind="1" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="2" <%if oitem.FOneItem.Fsourcekind="2" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 수산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="3" <%if oitem.FOneItem.Fsourcekind="3" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 축산물</span>
	  	<span style="margin-right:10px;"><input type="radio" name="rdArea" value="4" <%if oitem.FOneItem.Fsourcekind="4" then%>checked<%end if%> onClick="jsSetArea(this.value);"> 농수산가공품</span>
	  </p>
	  <p><input type="text" name="sourcearea" maxlength="64" size="64" class="text" id="[on,off,off,off][원산지]"  value="<%= oitem.FOneItem.Fsourcearea %>"/></p>
	  <div id="dvArea0" style="display:<%if isNull(oitem.FOneItem.Fsourcekind) or oitem.FOneItem.Fsourcekind="0" then%>block<%else%>none<%end if%>;">
	  <p><strong>ex: 한국, 중국, 중국OEM, 일본 등 </strong></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea1" style="display:<%if oitem.FOneItem.Fsourcekind ="1" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산, 국내산 또는 시·도명, 시·군명(대한민국, 한국X)  <span style="margin-right:10px;">ex. 쌀(국산)</span></BR>
	   <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 곶감(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea2" style="display:<%if oitem.FOneItem.Fsourcekind ="2" then%>block<%else%>none<%end if%>;">
	  <p><strong>국내산 :</strong> 국산,국내산 또는 연근해산(양식 수산물은 시·군명 가능)   <span style="margin-right:10px;">ex. 갈치(국산), 오징어(연근해산)</span> </BR>
	  	<strong>원양산 :</strong> 원양산 또는 원양산(해역명)   <span style="margin-right:10px;">ex. 참치[원양산(대서양)]</span> </BR>
	    <strong>수입산 :</strong> 통관시의 수입국가명 <span style="margin-right:10px;">ex. 농어(중국산)</span></BR>
	   - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea3" style="display:<%if oitem.FOneItem.Fsourcekind ="3" then%>block<%else%>none<%end if%>;">
	  <p>소고기의 경우 식육의 종류(한우/육우/젖소구분) 및 원산지   <span style="margin-right:10px;">ex. 쇠고기(횡성산 한우), 쇠고기(호주산)</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div>
	  <div id="dvArea4" style="display:<%if oitem.FOneItem.Fsourcekind ="4" then%>block<%else%>none<%end if%>;">
	  <p><strong>98%이상 원료가 있는 경우:</strong>  한가지 원료만 표시 가능    <span style="margin-right:10px;">ex. 쇠고기(미국산)</span> </BR>
	  	<strong>복합 원료를 사용한 경우:</strong> 혼합비율이 높은 순으로 2개 국가   <span style="margin-right:10px;">ex. 고추장[밀가루(미국산),고춧가루(국내산)]</span></BR>
	  - 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.</p>
	  </div> 
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][제조사]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(제조업체명)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="키워드 검색에서 사용될 추가 단어들" style="cursor:help;">검색키워드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="keywords" maxlength="250" size="50" class="text" id="[on,off,off,off][검색키워드]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="상품 상세 속성" style="cursor:help;">상품속성 :</td>
	<td id="lyrItemAttribAdd" bgcolor="#FFFFFF" colspan="3"></td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="upchemanagecode" class="text" id="[off,off,off,off][업체상품코드]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
		(업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">ISBN :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		ISBN 13 <input type="text" name="isbn13" class="text" value="<%= oitem.FOneItem.Fisbn13 %>" size="13" maxlength="13">
		/ 부가기호 <input type="text" name="isbn_sub" class="text" value="<%= oitem.FOneItem.FisbnSub %>" size="5" maxlength="5"><br />
		ISBN 10 <input type="text" name="isbn10" class="text" value="<%= oitem.FOneItem.Fisbn10 %>" size="10" maxlength="10"> (Optional)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">연관상품등록 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
	    <input type="text" name="relateItems" value="<%=strItemRelation%>" size="52" class="text" id="[off,off,off,off][연관상품]">
	    <br>(연관상품은 최대 6개까지 등록가능, 상품번호를 콤마(,)로 구분하여 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 설명 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<p style="text-align:right;"><input type="button" value="상품 이미지 보기" class="bgBlue" onClick="window.open('http://www.10x10.co.kr/shopping/itemImageView.asp?itemid=<%=itemid%>');"></p>
		<div>
		<!--
		<label><input type="radio" name="usinghtml" value="N" <% if oitem.FOneItem.Fusinghtml = "N" then response.write "checked" %>>일반TEXT</label>
		<label><input type="radio" name="usinghtml" value="H" <% if oitem.FOneItem.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML</label>
		<label><input type="radio" name="usinghtml" value="Y" <% if oitem.FOneItem.Fusinghtml = "Y" then response.write "checked" %>>HTML사용</label>
		<br>
		-->
		<input type="hidden" name="usinghtml" value="Y" />
		<textarea name="itemcontent" rows="15" class="textarea" style="width:100%" id="[on,off,off,off][아이템설명]"><%= oitem.FOneItem.Fitemcontent %></textarea>
		<script>
		//
		window.onload = new function(){
			var itemContEditor = CKEDITOR.replace('itemcontent',{
				height : 400,
				// 업로드된 파일 목록
				//filebrowserBrowseUrl : '/browser/browse.asp',
				// 파일 업로드 처리 페이지
				filebrowserImageUploadUrl : '<%= ItemUploadUrl %>/linkweb/items/itemEditorContentUpload.asp?itemid=<%=itemid%>'
			});
			itemContEditor.on( 'change', function( evt ) {
			    // 입력할 때 textarea 정보 갱신
			    document.itemreg.itemcontent.value = evt.editor.getData();
			});
		}
		</script>
		</div>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="80" id="[off,off,off,off][아이템동영상]"><%=oitemvideo.FOneItem.FvideoFullUrl%></textarea>
	    <br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="ordercomment" rows="5" cols="80" class="textarea" id="[off,off,off,off][유의사항]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
	<font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
	</td>
</tr>
</table>

<!-- 품목상세정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">품목상세정보 &nbsp;<font color=gray>상품정보제공고시 관련 법안 추진에 따라 아래 내용을 정확히 입력해주시기 바랍니다.</font></td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목선택 :</td>
	<td bgcolor="#FFFFFF">
		<% DrawInfoDiv "infoDiv", oitem.FOneItem.FinfoDiv, " onchange='chgInfoDiv(this.value);'" %>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
	</td>
</tr>
</table>
<!-- 안전인증정보 -->
<%
dim arrAuth, r, real_safetydiv, real_safetynum, safetyDivList
arrAuth = oitem.FAuthInfo
if isArray(arrAuth) THEN
	For r =0 To UBound(arrAuth,2)
		real_safetydiv = real_safetydiv & arrAuth(0,r)
		if r <> UBound(arrAuth,2) then real_safetydiv = real_safetydiv & "," end if
		
		real_safetynum = real_safetynum & arrAuth(1,r)
		if r <> UBound(arrAuth,2) then real_safetynum = real_safetynum & "," end if
		
		safetyDivList = safetyDivList & "<p class='tPad05' id='l"&arrAuth(0,r)&"'>"
		safetyDivList = safetyDivList & "- "&fnSafetyDivCodeName(arrAuth(0,r))&"("&CHKIIF(arrAuth(1,r)="x","인증번호 없음",arrAuth(1,r))&")"
		safetyDivList = safetyDivList & " <input type='button' value='삭제' class='btn3 btnIntb' onClick='jsSafetyDivListDel("&arrAuth(0,r)&");'>"
		safetyDivList = safetyDivList & "</p>"
	Next
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">안전인증정보</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">
		안전인증대상 :
		<input type="button" value="안전인증 필수 품목 확인" onclick="jsSafetyPopup();" class="button" />
	</td>
	<td bgcolor="#FFFFFF">
		<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr align="left" height="30">
			<td bgcolor="#FFFFFF">
				<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상</label>
				<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn="N","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 대상아님</label>
				<label><input type="radio" name="safetyYn" value="I" <%=chkIIF(oitem.FOneItem.FsafetyYn="I","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 상품설명에 표기</label>
				<label><input type="radio" name="safetyYn" value="S" <%=chkIIF(oitem.FOneItem.FsafetyYn="S","checked","")%> onclick="chgSafetyYn(document.itemreg)" /> 안전기준준수</label>
				<input type="hidden" name="auth_go_catecode" id="auth_go_catecode" value="">
				<input type="hidden" name="real_safetydiv" id="real_safetydiv" value="<%=real_safetydiv%>">
				<input type="hidden" name="real_safetynum" id="real_safetynum" value="<%=real_safetynum%>">
				<input type="hidden" name="real_safetyidx" id="real_safetyidx" value="">
				<input type="hidden" name="real_safetynum_delete" id="real_safetynum_delete" value="">
				<input type="hidden" name="real_safetydiv_delete" id="real_safetydiv_delete" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<% drawSelectBoxSafetyDivCode "safetyDiv", "", oitem.FOneItem.FsafetyYn, "" %>

				인증번호 <input type="text" name="safetyNum" id="[off,off,off,off][안전인증 인증번호]" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" value="" /><%'=oitem.FOneItem.FsafetyNum%>
				<input type="button" id="safetybtn" value="추   가" onclick="jsSafetyAuth();" class="button">
				<input type="hidden" name="issafetyauth" id="issafetyauth" value="">
			</td>
		</tr>
		<tr align="left">
			<td bgcolor="#FFFFFF">
				<div id="safetyDivList">
					<%=safetyDivList%>
				</div>
				<div id="safetyYnI" style="display:none;">
					<font color="blue">상품 설명에 표기(표기대상 상품인경우 상품 상세 페이지에 인증번호와 모델명, KC 마크를 꼭 표기해주세요.)</font>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="left">
	<td bgcolor="#FFFFFF" colspan=2>
		* 인증정보를 입력 안 하거나, 잘못된 인증정보를 입력한 경우 발견 <strong><font color='red'>즉시 판매정지 또는 삭제</font></strong> 됩니다.<br>
		* <strong><font color='red'>안전기준준수</font></strong> 대상일경우 인증번호가 없으며, KC마크를 표시하지 않아야 됩니다.<br>
		* 입력한 인증정보는 제품안전정보센터에서 제공된 정보를 기준으로 조회되며, <strong><font color='red'>검증되지 않은 정보는 등록이 불가</font></strong>능합니다.<br>
		* 정상적인 인증정보를 입력했음에도 불구하고 등록이 안될경우에 "상품설명에 표기"로 설정이 가능하며, 상품 상세 페이지에 모델명과 표기대상 상품인경우 인증번호,KC마크를 표기해야 합니다.<br>
		* 안전인증정보 관련 문의는 홈페이지(<u><a href="http://safetykorea.kr" target="_blank">http://safetykorea.kr</a></u>)로 확인해 주시기 바랍니다.
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" class="button" onClick="SubmitSave()">
          <input type="button" value="취소하기" class="button" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
</form>

<script type="text/javascript">
	itemreg.makerid.readOnly = true;
	itemreg.frontMakerid.readOnly = true;

	// 안전인증체크. 전안법
	jsSafetyCheck('<%= oitem.FOneItem.FsafetyYn %>','');
</script>

<% 
set oitem = Nothing
Set oitemvideo = Nothing

Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- 업체선택 --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c order by userid"
'   query1 = query1 + " where isusing='Y' order by userid desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Lcase(selectedId) = Lcase(rsget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("userid")& "," & rsget("defaultmargine") & "' "&tmp_str&">" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")
End Sub
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->