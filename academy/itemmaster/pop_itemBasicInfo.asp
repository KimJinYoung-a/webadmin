<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%

dim itemid, oitem , oitemvideo
dim makerid
Dim fingerson : fingerson = "on" '//상품고시용 fingersflag

itemid = RequestCheckvar(request("itemid"),10)
makerid = RequestCheckvar(request("makerid"),32)
menupos = RequestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

Set oitemvideo = New CItem
oitemvideo.FRectItemId = itemid
oitemvideo.FRectItemVideoGubun = "video1"
oitemvideo.GetItemContentsVideo

'==============================================================================
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

'==============================================================================
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
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript" SRC="/js/confirm.js"></script>
<script>
function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 카테고리등록
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

	popwin = window.open('/academy/comm/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
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
        alert('주문제작 상품은 업체배송인경우만 가능합니다.');
        itemreg.itemdiv[0].focus();
        return;
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

	//안전인증정보
    if (itemreg.safetyYn[0].checked){
	    if (!itemreg.safetyDiv.value){
	        alert('안전인증구분을 선택해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
	    if (!itemreg.safetyNum.value){
	        alert('안전인증번호를 입력해주세요.');
	        itemreg.safetyDiv.focus();
	        return;
	    }
    }

    if(confirm("상품을 올리시겠습니까?") == true){
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
		var popwin = window.open("/academy/comm/NewCategorySelect.asp?iid=" + iid, "popCateSel","width=700,height=400,scrollbars=yes,resizable=yes");
        popwin.focus();
	}

	// 전시카테고리 선택 팝업
	function popDispCateSelect(){
		var designerid = document.all.itemreg.designerid.value;
		if(designerid == ""){
			alert("업체를 선택하세요.");
			return;
		}
		
		var dCnt = $("input[name='isDefault'][value='y']").length;
		$.ajax({
			url: "/academy/comm/act_DispCategorySelect.asp?designerid="+designerid+"&isDft="+dCnt,
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
		//printItemAttribute();
	}

	// 선택 카테고리 삭제
	function delCateItem()
	{
		if(confirm("선택한 카테고리를 삭제하시겠습니까?"))
			tbl_Category.deleteRow(tbl_Category.clickedRowIndex);
	}
	
	// 선택 전시카테고리 삭제
	function delDispCateItem() {
		if(confirm("선택한 카테고리를 삭제하시겠습니까?")) {
			tbl_DispCate.deleteRow(tbl_DispCate.clickedRowIndex);

			//상품속성 출력
			//printItemAttribute();
		}
	}
//-------------------------------------
function chgodr(v){
	if (v == 1){
		$("#customorder").css("display","none");
	}else{
		$("#customorder").css("display","");
	}
}

function chgodr2(v){
	if (v == 1){
		$("#subodr").css("display","none");
	}else{
		$("#subodr").css("display","");
	}
}

// 안전인증정보 선택
function chgSafetyYn(frm) {
	if(frm.safetyYn[0].checked) {
		frm.safetyDiv.disabled=false;
		frm.safetyNum.disabled=false;
	} else {
		frm.safetyDiv.disabled=true;
		frm.safetyNum.disabled=true;
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
			url: "/admin/itemmaster/act_itemInfoDivForm.asp",
			data: "itemid=<%=itemid%>&ifdv="+v+"&fingerson=on",
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

function checkItemDiv(comp){
    var frm = comp.form;
    
    if (comp.name=="itemdiv"){
        if (frm.itemdiv[1].checked){
            frm.reqMsg.disabled=false;
            frm.requireimgchk.disabled=false;
        }else{
            //frm.reqMsg.checked=false;
            frm.reqMsg.disabled=true;
            frm.requireimgchk.disabled=true;
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
}

function requireimg(){
	var frm = document.itemreg;
	if (frm.requireimgchk.checked){
		$("#rmemail").css("display","");
	}else{
		$("#rmemail").css("display","none");
	}
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
        <td align="left">
          <br>기본정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;" style="margin:0;">
<input type="hidden" name="mode" value="ItemBasicInfo">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">
<input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
<input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= itemid %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="미리보기" onclick="window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= itemid %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">브랜드ID :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fmakerid %></td>
  </tr>
  <!-- // 멀티 카테고리 추가로인해 제거 (업체 상품수정에서는 기본카테고리 지정용으로 사용) (2008.03.28;허진원 수정)
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">카테고리 구분 :</td>
    <input type="hidden" name="cd1" value="<%= oitem.FOneItem.FCate_large %>">
    <input type="hidden" name="cd2" value="<%= oitem.FOneItem.FCate_mid %>">
    <input type="hidden" name="cd3" value="<%= oitem.FOneItem.FCate_small %>">
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="cd1_name" value="<%= oitem.FOneItem.FCate_large_name %>" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="<%= oitem.FOneItem.FCate_mid_name %>" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="<%= oitem.FOneItem.FCate_small_name %>" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">

      <input type="button" value="카테고리 선택" onclick="editCategory(itemreg.cd1.value,itemreg.cd2.value,itemreg.cd3.value);">
  	</td>
  </tr>
  //-->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">관리 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getCategoryInfo(itemid)%></td>
			<td valign="bottom"><input type="button" value="추가" onClick="popCateSelect('<%=itemid%>')"></td>
		</tr>
		</table>
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<table class=a>
		<tr>
			<td><%=getDispCategory(itemid)%></td>
			<td valign="bottom"><input type="button" value="+" class="button" onClick="popDispCateSelect()"></td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="itemdiv" value="01" <% if oitem.FOneItem.Fitemdiv ="01" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr2(1);">일반상품
      <input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onclick="checkItemDiv(this);chgodr(2);">주문제작상품
      <input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문제작 메세지가 필요한 경우)</font>
	  <input type="checkbox" name="requireimgchk" value="Y" <%=chkIIF(oitem.FOneItem.Frequirechk="Y","checked","")%> onClick="requireimg();">주문제작 이미지 필요
<!-- 	  <br> -->
<!--       <input type="radio" name="itemdiv" value="20" <% if oitem.FOneItem.Fitemdiv ="20" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">추가전용상품 -->
<!--       <font color="red">(상품목록에서는 제외, 추가옵션에서만 보여짐)</font> -->
      <% if oitem.FOneItem.Fitemdiv ="07" then %>
      <input type="radio" name="itemdiv" value="07" <% if oitem.FOneItem.Fitemdiv ="07" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">공동(예약)구매상품
      <% end if %>
      
      <% if oitem.FOneItem.Fitemdiv ="82" then %>
      <input type="radio" name="itemdiv" value="82" <% if oitem.FOneItem.Fitemdiv ="82" then  response.write "checked" %> onclick="checkItemDiv(this);chgodr(1);">마일리지샵 상품
      <% end if %>
  	</td>
  </tr>
    <!-- 주문 제작 이메일 -->
  <tr id="rmemail" style="display:<%=chkiif(oitem.FOneItem.Frequirechk="Y","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 이메일 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="requireMakeEmail" value="<%=oitem.FOneItem.FrequireEmail%>" size="50" maxlength="100"> (ex)작가님의 메일 주소)
  	</td>
  </tr>
  <!-- 주문 제작 이메일 -->
  <tr id="customorder" style="display:<%=chkiif(oitem.FOneItem.Fitemdiv="06" Or oitem.FOneItem.Fitemdiv="16","","none")%>;">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문제작 추가옵션</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="radio" name="cstodr" value="1" onclick="chgodr2(1)" <%=chkiif(oitem.FOneItem.Fcstodr="1","checked","")%>>즉시발송
      <input type="radio" name="cstodr" value="2" onclick="chgodr2(2)" <%=chkiif(oitem.FOneItem.Fcstodr="2","checked","")%>>제작후 발송<br>
	  <div id="subodr" style="display:<%=chkiif(oitem.FOneItem.Fcstodr="2","block","none")%>;">
		제작후 발송 기간 <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="3" maxlength="2">일<br>
		&lt--특이사항을 입력 해주세요--&gt;<br><textarea name="requirecontents" rows="5" cols="80"><%=oitem.FOneItem.Frequirecontents%></textarea>
	  </div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemname" maxlength="64" size="60" id="[on,off,off,off][상품명]" value="<%= oitem.FOneItem.Fitemname %>">&nbsp;
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsource" maxlength="64" size="50" id="[on,off,off,off][상품재질]" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemsize" maxlength="64" size="50" id="[on,off,off,off][상품사이즈]" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품무게 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="itemWeight" maxlength="12" size="4" id="[on,off,off,off][상품무게]" value="<%= oitem.FOneItem.FitemWeight %>">g&nbsp;(무게는 g단위로 입력)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">원산지 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="sourcearea" maxlength="64" size="25" id="[on,off,off,off][원산지]" value="<%= oitem.FOneItem.Fsourcearea %>">&nbsp;(ex:한국,중국,중국OEM,일본...)
      <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">제조사 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="makername" maxlength="32" size="25" id="[on,off,off,off][제조사]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(제조업체명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">검색키워드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <input type="text" name="keywords" maxlength="50" size="50" id="[on,off,off,off][검색키워드]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	    <input type="text" name="upchemanagecode" id="[off,off,off,off][업체상품코드]" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="20" maxlength="32">
  	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="radio" name="usinghtml" value="N" <% if oitem.FOneItem.Fusinghtml = "N" then response.write "checked" %>>일반TEXT -->
<!--       <input type="radio" name="usinghtml" value="H" <% if oitem.FOneItem.Fusinghtml = "H" then response.write "checked" %>>TEXT+HTML -->
<!--       <input type="radio" name="usinghtml" value="Y" <% if oitem.FOneItem.Fusinghtml = "Y" then response.write "checked" %>>HTML사용 -->
<!--       <br> -->
<!--       <textarea name="itemcontent" rows="10" cols="80" id="[on,off,off,off][아이템설명]"><%= oitem.FOneItem.Fitemcontent %></textarea> -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :<br/>(배송안내)</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="ordercomment" rows="5" cols="80" id="[off,off,off,off][유의사항]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
      <font color="red">특별한 배송기간이나 주문시 확인해야만 하는 사항</font>을 입력하시면 고객불만이나 환불을 줄일수 있습니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">교환 / 환불 정책</td>
  	<td bgcolor="#FFFFFF" colspan="3">
      <textarea name="refundpolicy" rows="5" cols="80" id="[off,off,off,off][환불정책]"><%=oitem.FOneItem.Frefundpolicy%></textarea><br>
  	</td>
  </tr>
<!--   <tr align="left"> -->
<!--   	<td height="30" width="15%" bgcolor="#DDDDFF">업체코멘트 :</td> -->
<!--   	<td bgcolor="#FFFFFF" colspan="3"> -->
<!--       <input type="text" name="designercomment" size="50" maxlength="40" id="[off,off,off,off][업체코멘트]" value="<%= oitem.FOneItem.Fdesignercomment %>"><br> -->
<!--       상품에관한 스토리나 재미난 이야기를 적어주세요... -->
<!--   	</td> -->
<!--   </tr> -->
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">아이템 동영상 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="itemvideo" rows="5" class="textarea" cols="90" id="[off,off,off,off][아이템동영상]"><%= db2html(oitemvideo.FOneItem.FvideoFullUrl) %></textarea>
		<br>※ Youtube, Vimeo 동영상만 가능(Youtube : 소스코드값 입력, Vimeo : 임베딩값 입력)
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
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::상품품목::</option>
		<option value="01">의류</option>
		<option value="02">구두/신발</option>
		<option value="03">가방</option>
		<option value="04">패션잡화(모자/벨트/액세서리)</option>
		<option value="05">침구류/커튼</option>
		<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
<!-- 		<option value="07">영상가전(TV류)</option> -->
<!-- 		<option value="08">가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option> -->
<!-- 		<option value="09">계절가전(에어컨/온풍기)</option> -->
<!-- 		<option value="10">사무용기기(컴퓨터/노트북/프린터)</option> -->
<!-- 		<option value="11">광학기기(디지털카메라/캠코더)</option> -->
<!-- 		<option value="12">소형전자(MP3/전자사전 등)</option> -->
<!-- 		<option value="14">내비게이션</option> -->
		<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
<!-- 		<option value="16">의료기기</option> -->
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
<!-- 		<option value="27">호텔/펜션예약</option> -->
<!-- 		<option value="28">여행상품</option> -->
<!-- 		<option value="29">항공권</option> -->
		<option value="35">기타</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitem.FOneItem.FinfoDiv%>";
		chgInfoDiv(<%=oitem.FOneItem.FinfoDiv%>);
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") Then
			Server.Execute("/admin/itemmaster/act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<!-- <tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품재질 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...) -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;"> -->
<!-- 	<td height="30" width="15%" bgcolor="#DDDDFF">상품사이즈 :</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm)) -->
<!-- 	</td> -->
<!-- </tr> -->
</table>
<!-- 안전인증정보 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">안전인증정보</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#F8DDFF">안전인증대상 :</td>
	<td bgcolor="#FFFFFF">
		<label><input type="radio" name="safetyYn" value="Y" <%=chkIIF(oitem.FOneItem.FsafetyYn="Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> 대상</label>
		<label><input type="radio" name="safetyYn" value="N" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","checked","")%> onclick="chgSafetyYn(document.itemreg)"> 대상아님</label> /
		<select name="safetyDiv" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> class="select">
		<option value="">::안전인증구분::</option>
		<option value="10" <%=chkIIF(oitem.FOneItem.FsafetyDiv="10","selected","")%>>국가통합인증(KC마크)</option>
		<option value="20" <%=chkIIF(oitem.FOneItem.FsafetyDiv="20","selected","")%>>전기용품 안전인증</option>
		<option value="30" <%=chkIIF(oitem.FOneItem.FsafetyDiv="30","selected","")%>>KPS 안전인증 표시</option>
		<option value="40" <%=chkIIF(oitem.FOneItem.FsafetyDiv="40","selected","")%>>KPS 자율안전 확인 표시</option>
		<option value="50" <%=chkIIF(oitem.FOneItem.FsafetyDiv="50","selected","")%>>KPS 어린이 보호포장 표시</option>
		</select>
		인증번호 <input type="text" name="safetyNum" <%=chkIIF(oitem.FOneItem.FsafetyYn<>"Y","disabled","")%> size="35" maxlength="25" class="text" value="<%=oitem.FOneItem.FsafetyNum%>" />
		
		<font color="darkred">유아용품이나 전기용품일 경우 필수 입력</font>
	</td>
</tr>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" onClick="SubmitSave()">
          <input type="button" value="취소하기" onClick="self.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
</form>
<!-- 표 하단바 끝-->
<% 
set oitem = Nothing
Set oitemvideo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->