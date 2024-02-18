<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%

dim itemid, oitem
dim makerid

itemid = requestCheckVar(request("itemid"),20)
makerid = requestCheckVar(request("makerid"),50)
menupos = requestCheckVar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.FRectMakerId = session("ssBctID")
oitem.GetOneItem

if (oitem.FResultCount < 1) then
    response.write "<script>alert('잘못된 접속입니다..');</script>"
    dbget.close()	:	response.End
end if


'==============================================================================
'세일마진
dim sailmargine
'가격계산
if oitem.FOneItem.Fsailyn="Y" then
	 if oitem.FOneItem.Fvatinclude = "Y" then
			on error resume next
			sailmargine = fix((CLng(oitem.FOneItem.Fsailprice)-Clng(oitem.FOneItem.Fsailsuplycash))/CLng(oitem.FOneItem.Fsailprice)*100*100)/100
			if Err then
				sailmargine = 0
			end if
	 else
			on error resume next
			sailmargine = fix((CLng(oitem.FOneItem.Fsailprice)-Clng(oitem.FOneItem.Fsailsuplycash)-CLng(oitem.FOneItem.Fbuyvat))/CLng(oitem.FOneItem.Fsailprice)*100*100)/100
			if Err then
				sailmargine = 0
			end if
	 end if
else
    sailmargine = 0
end if


'==============================================================================
Sub SelectBoxDesignerItem(selectedId)
   dim query1,tmp_str
   %><select name="designer" onchange="TnDesignerNMargineAppl(this.value);">
     <option value='' <%if selectedId="" then response.write " selected"%>>-- 업체선택 --</option><%
   query1 = " select userid,socname_kor,defaultmargine from [db_user].[dbo].tbl_user_c"
'   query1 = query1 + " where isusing='Y'"
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
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
// ============================================================================
// 카테고리등록(사용안함;2010-09-13 허진원-MD요청에 의해 삭제)
/*
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
*/

function popMultiLangEdit(iid) {
	window.open("/common/item/pop_MultiLangItemCont.asp?itemid="+iid+"&lang=EN", "multiLang_win", "width=1280, height=960, scrollbars=yes, resizable=yes");
}


// ============================================================================
// 저장하기
function SubmitSave() {
    if (validate(itemreg)==false) {
        return;
    }
    //업체배송만 주문제작 가능.
    <% if oitem.FOneItem.Fmwdiv <> "U" then %>
   if(typeof(itemreg.itemdiv.length)!="undefined"){ 
	    if (itemreg.itemdiv[1].checked){
	        alert('주문제작 상품은 업체배송인경우만 가능합니다.');
	        itemreg.itemdiv[0].focus();
	        return;
	    }
	  }
    <% end if %>

	//상품 설명에 불가항목 검사
	var cntRe = /.js["'>\s]/gi;
	if(cntRe.test(itemreg.itemcontent.value)) {
        alert('상품설명에는 js파일을 넣을 수 없습니다.');
        itemreg.itemcontent.focus();
        return;
	}
	
	//상품무게 숫자체크
 if (!IsDigit(itemreg.itemWeight.value)){
		alert('상품무게는  숫자로 입력하세요.');
		itemreg.itemWeight.focus();
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

//해외배송
		if(document.itemreg.optionaddprice.value >0 && document.itemreg.deliverOverseas.checked){
			alert("옵션에 추가가격이 있을 경우 해외배송이 불가능합니다. 해외배송체크를 해제해주세요" );
			document.itemreg.deliverOverseas.focus();
			 return;
		}
		
		
 	if(document.itemreg.deliverOverseas.checked){
	    if(document.itemreg.itemWeight.value<=0){
	        alert("해외배송시 배송비 산출을 위해 상품무게를 꼭 입력해주세요")
	        document.itemreg.itemWeight.focus();
	        return;
	    }
	} 

	//화물반송비
	try{
		if(itemreg.freight_min.value<=0||itemreg.freight_max.value<=0) {
            alert('화물배송 비용을 입력해주세요.');
            itemreg.freight_min.focus();
            return;
		}
	} catch(e) {}

    if(confirm("상품을 올리시겠습니까?") == true){
        itemreg.submit();
    }

}

function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp','op2','width=450,height=570,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ClearVal(comp){
    comp.value = "";
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
			<font color="red"><strong>상품수정</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			등록된 상품을 수정합니다.<br>
			문의사항이 있으신 분은 각 카테고리별 MD에게 문의하시면 됩니다.
			&nbsp;&nbsp;
			<a href="javascript:pop_10x10_person();"><img src="/images/icon_arrow_link.gif" border="0" align="absbottom">&nbsp;카테고리별 MD연락처</a> 
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
<form name="itemreg" method="post" action="do_upche_item_infomodify.asp" onsubmit="return false;">
<input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
<input type="hidden" name="optionaddprice" value="<%=oitem.FOneItem.fnGetOptAddPrice(oitem.FOneItem.Fitemid)%>">
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top"><img src="/images/icon_arrow_down.gif" border="0" align="absbottom">
	        	<strong>기본정보</strong></td>
	        <td valign="top" align="right">&nbsp;</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->



<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
  	  <%= oitem.FOneItem.Fitemid %>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">업체명 :</td>
  	<td bgcolor="#FFFFFF" colspan="2"><%= oitem.FOneItem.Fmakerid %></td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">카테고리 구분 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="hidden" name="cd1" value="<%= oitem.FOneItem.FCate_large %>">
      <input type="hidden" name="cd2" value="<%= oitem.FOneItem.FCate_mid %>">
      <input type="hidden" name="cd3" value="<%= oitem.FOneItem.FCate_small %>">
      <input type="text" name="cd1_name" value="<%= oitem.FOneItem.FCate_large_name %>" class="text" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="<%= oitem.FOneItem.FCate_mid_name %>" class="text" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="<%= oitem.FOneItem.FCate_small_name %>" class="text" id="[on,off,off,off][카테고리]" size="20" readonly style="background-color:#E6E6E6">
  	</td>
  </tr>
  <tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF" title="프론트에 진열될 카테고리" style="cursor:help;">전시 카테고리 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table class=a>
		<tr>
			<td id="lyrDispList"><%=getDispOnlyCategory(oitem.FOneItem.Fitemid)%></td>
			<td valign="bottom"> </td>
		</tr>
		</table>
		<div id="lyrDispCateAdd" style="border:1px solid #CCCCCC; border-radius: 6px; background-color:#F8F8FF; padding:6px; display:none;"></div>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">상품구분 :</td>
  	<td bgcolor="#FFFFFF" >
      <% if oitem.FOneItem.Fitemdiv="08" then %> 	 
     					<input type="radio" name="itemdiv" value="08" <%=chkIIF(oitem.FOneItem.Fitemdiv="08","checked","")%>  >티켓상품 
     				<% elseif oitem.FOneItem.Fitemdiv="09" then %> 	 
		 				<input type="radio" name="itemdiv" value="09" <%=chkIIF(oitem.FOneItem.Fitemdiv="09","checked","")%> >Present상품 
		 			<% elseif oitem.FOneItem.Fitemdiv="18" then %> 	 	
	 					<input type="radio" name="itemdiv" value="18" <%=chkIIF(oitem.FOneItem.Fitemdiv="18","checked","")%>  >여행상품  
					<% elseif oitem.FOneItem.Fitemdiv ="82" then %>
			        <input type="radio" name="itemdiv" value="82" <%=chkIIF(oitem.FOneItem.Fitemdiv="82","checked","")%>  >마일리지샵 상품 
		 			<% elseif oitem.FOneItem.Fitemdiv ="75" then %> 
		 			<input type="radio" name="itemdiv" value="75" <%=chkIIF(oitem.FOneItem.Fitemdiv="75","checked","")%> >정기구독상품 
		 				<% elseif oitem.FOneItem.Fitemdiv ="07" then %> 
		 			<input type="radio" name="itemdiv" value="07" <%=chkIIF(oitem.FOneItem.Fitemdiv="07","checked","")%> >구매제한상품 
      <% else %>
		<label><input type="radio" name="itemdiv" value="01" <%=chkIIF(oitem.FOneItem.Fitemdiv ="01","checked","")%> onClick="this.form.requireMakeDay.value=0;document.getElementById('lyRequre').style.display='none';checkItemDiv(this);">일반상품</label>
		<br>
		<label><input type="radio" name="itemdiv" value="<%= oitem.FOneItem.Fitemdiv %>" <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","checked","")%> onClick="document.getElementById('lyRequre').style.display='block';checkItemDiv(this);">주문 제작상품</label>
		<input type="checkbox" name="reqMsg" value="10" <%=chkIIF(oitem.FOneItem.Fitemdiv="06","checked","")%> <%=chkIIF(oitem.FOneItem.Fitemdiv="06" or oitem.FOneItem.Fitemdiv="16","","disabled")%> onClick="checkItemDiv(this);">주문제작 문구 필요<font color=red>(주문시 이니셜등 제작문구가 필요한경우 체크)</font>
		<br>
		<%if not (oitem.FOneItem.Fitemdiv ="01" or oitem.FOneItem.Fitemdiv ="06" or oitem.FOneItem.Fitemdiv ="16") then%>
							<input type="hidden" name="itemdiv" id="itemdiv" value="<%=oitem.FOneItem.Fitemdiv%>">
							<%end if%>
      <% end if %>
  	</td>
  	<td bgcolor="#FFFFFF" >
  	    <div id="lyRequre" style="<%=chkIIF(oitem.FOneItem.Fitemdiv ="06" or oitem.FOneItem.Fitemdiv ="16","","display:none;")%>padding-left:22px;">
			예상제작소요일 <input type="text" name="requireMakeDay" value="<%=oitem.FOneItem.FrequireMakeDay%>" size="2" class="text" id="[off,on,off,off][예상제작소요일]">일
			<font color="red">(상품발송전 상품제작 기간)</font>
		</div>
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <%= oitem.FOneItem.Fitemname %>&nbsp;
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">영문상품명 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" name="itemnameEng" maxlength="64" size="60" class="text_ro" readonly id="[off,off,off,off][영문상품명]" value="<%= oitem.FOneItem.FitemnameEng %>">&nbsp;
		<input type="button" value="영문 정보 <%=chkIIF(oitem.FOneItem.FitemnameEng="" or isnull(oitem.FOneItem.FitemnameEng),"등록","수정")%>" class="button" onclick="popMultiLangEdit(<%= oitem.FOneItem.Fitemid %>)" />
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">원산지 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="sourcearea" maxlength="64" size="25" class="text" id="[on,off,off,off][원산지]" value="<%= oitem.FOneItem.Fsourcearea %>">&nbsp;(ex:한국,중국,중국OEM,일본 등 / 식품일 경우 국내: 국내산 또는 시군구명, 수입: 미국산, 중국산 등)
      <br>( 원산지 표기 오류는 고객클레임의 가장 큰 원인 중 하나입니다. 정확히 입력해 주세요.)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">제조사 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="makername" maxlength="32" size="25" class="text" id="[on,off,off,off][제조사]" value="<%= oitem.FOneItem.Fmakername %>">&nbsp;(제조업체명)
  	</td>
  </tr>
  <tr align="left">
	<td height="30" width="160" bgcolor="#DDDDFF">상품무게 :</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="itemWeight" maxlength="12" size="8" id="[on,off,off,off][상품무게]" style="text-align:right" value="<%= oitem.FOneItem.Fitemweight %>">g &nbsp;(그램단위로 입력, ex:1.5kg→ 1500) / 해외배송시 배송비 산출을 위한 것이므로 정확히 입력.
	</td>
</tr>
<tr align="left">
	<td height="30" width="160" bgcolor="#DDDDFF">배송지역 :</td>
	<td   bgcolor="#FFFFFF" colspan="3">
	  <input type="radio" name="deliverarea" value="" <%=chkIIF(Trim( oitem.FOneItem.Fdeliverarea)="" or IsNull( oitem.FOneItem.Fdeliverarea),"checked","")%>>전국배송&nbsp;
	  <input type="radio" name="deliverarea" value="C" <%=chkIIF( oitem.FOneItem.Fdeliverarea="C","checked","")%> <%if oitem.FOneItem.Fdeliverfixday<>"C" then%>disabled<%end if%>>수도권배송&nbsp;
	  <input type="radio" name="deliverarea" value="S" <%=chkIIF( oitem.FOneItem.Fdeliverarea="S","checked","")%> <%if oitem.FOneItem.Fdeliverfixday<>"C" then%>disabled<%end if%>>서울배송&nbsp;
	  <label><input type="checkbox" name="deliverOverseas" value="Y" <%=chkIIF( oitem.FOneItem.FdeliverOverseas="Y","checked","")%> title="해외배송은 상품무게가 입력이 돼야 완료됩니다.">해외배송</label>
	</td>
</tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">검색키워드 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
      <input type="text" name="keywords" maxlength="260" size="120" class="text" id="[on,off,off,off][검색키워드]" value="<%= oitem.FOneItem.Fkeywords %>">&nbsp;(콤마로구분 ex: 커플,티셔츠,조명)
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="160" bgcolor="#DDDDFF">업체상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="2">
  	    <input type="text" name="upchemanagecode" class="text" value="<%= oitem.FOneItem.Fupchemanagecode %>" size="30" maxlength="32" id="[off,off,off,off][업체상품코드]">
  	    (업체에서 관리하는 코드 최대 32자 - 영문/숫자만 가능)
  	</td>
  </tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">상품 설명 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <input type="radio" name="usinghtml" value="N" <%=chkIIF(oitem.FOneItem.Fusinghtml="N","checked","")%>>일반TEXT
	  <input type="radio" name="usinghtml" value="H" <%=chkIIF(oitem.FOneItem.Fusinghtml="H","checked","")%>>TEXT+HTML
	  <input type="radio" name="usinghtml" value="Y" <%=chkIIF(oitem.FOneItem.Fusinghtml="Y","checked","")%>>HTML사용
	  <br>
	  <textarea name="itemcontent" rows="15" class="textarea" style="width:100%" id="[on,off,off,off][상품설명]"><%= oitem.FOneItem.Fitemcontent %></textarea>
	</td>
</tr>
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">주문시 유의사항 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <textarea name="ordercomment" rows="5" cols="90" class="textarea" id="[off,off,off,off][유의사항]"><%= oitem.FOneItem.Fordercomment %></textarea><br>
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
		<select name="infoDiv" class="select" onchange="chgInfoDiv(this.value)">
		<option value="">::상품품목::</option>
		<option value="01">의류</option>
		<option value="02">구두/신발</option>
		<option value="03">가방</option>
		<option value="04">패션잡화(모자/벨트/액세서리)</option>
		<option value="05">침구류/커튼</option>
		<option value="06">가구(침대/소파/싱크대/DIY제품)</option>
		<option value="07">영상가전(TV류)</option>
		<option value="08">가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지)</option>
		<option value="09">계절가전(에어컨/온풍기)</option>
		<option value="10">사무용기기(컴퓨터/노트북/프린터)</option>
		<option value="11">광학기기(디지털카메라/캠코더)</option>
		<option value="12">소형전자(MP3/전자사전 등)</option>
		<option value="14">내비게이션</option>
		<option value="15">자동차용품(자동차부품/기타 자동차용품)</option>
		<option value="16">의료기기</option>
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
		<option value="35">기타</option>
		</select>
		<script type="text/javascript">
		document.itemreg.infoDiv.value="<%=oitem.FOneItem.FinfoDiv%>";
		</script>
	</td>
</tr>
<tr align="left" id="itemInfoCont" style="display:<%=chkIIF(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="","none","")%>;">
	<td height="30" width="15%" bgcolor="#F8DDFF">품목내용 :</td>
	<td bgcolor="#FFFFFF" id="itemInfoList">
	<%
		if Not(isNull(oitem.FOneItem.FinfoDiv) or oitem.FOneItem.FinfoDiv="") then
			Server.Execute("/admin/itemmaster/act_itemInfoDivForm.asp")
		end if
	%>
	</td>
</tr>
<tr align="left">
	<td height="25" colspan="2" bgcolor="#FDFDFD"><font color="darkred">상품상세페이지에 내용이 포함 되어있더라도 정확히 입력바랍니다. 부정확하거나 잘못된 정보 입력시, 그에 대한 책임을 물을 수도 있습니다.</font></td>
</tr>
<tr align="left" id="lyItemSrc" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
  	<td height="30" width="160" bgcolor="#DDDDFF">상품재질 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsource" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsource %>">&nbsp;(ex:플라스틱,비즈,금,...)
  	</td>
</tr>
<tr align="left" id="lyItemSize" style="display:<%=chkIIF(oitem.FOneItem.FinfoDiv="35","","none")%>;">
  	<td height="30" width="160" bgcolor="#DDDDFF">상품사이즈 :</td>
  	<td bgcolor="#FFFFFF">
      <input type="text" name="itemsize" maxlength="64" size="50" class="text" value="<%= oitem.FOneItem.Fitemsize %>">&nbsp;(ex:7.5x15(cm))
  	</td>
</tr>
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

<%
	'화물배송 반송비 입력 (화물배송일 때만)
	if oitem.FOneItem.Fdeliverfixday="X" then
%>
<!-- 화물배송 반송비 입력 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="left">화물배송 정보</td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td height="30" width="15%" bgcolor="#DDDDFF">화물배송 반송 비용 :</td>
	<td bgcolor="#FFFFFF" colspan="2">
		&nbsp;
		최소 <input type="text" name="freight_min" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_min%>" style="text-align:right;">원 ~
		최대 <input type="text" name="freight_max" class="text" size="6" value="<%=oitem.FOneItem.Ffreight_max%>" style="text-align:right;">원
		<br>&nbsp; <font color="red">(반품/교환 시 편도 비용)</font>
	</td>
</tr>
</table>
<%	end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="30">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <input type="button" value="저장하기" class="button" onClick="SubmitSave()">
          <input type="button" value="창 닫 기" class="button" onClick="window.close()">
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
<p>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->