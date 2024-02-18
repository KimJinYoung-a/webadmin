<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인상품 등록
' Hieditor : 2009.04.07 서동석 생성
'			 2010.06.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim itemgubun,itemid, itemoption, barcode ,i
	barcode	  = requestCheckVar(Trim(request("barcode")),32)

if BF_IsMaybeTenBarcode(barcode) then
    itemgubun 	= BF_GetItemGubun(barcode)
	itemid 		= BF_GetItemId(barcode)
	itemoption 	= BF_GetItemOption(barcode)
end if

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FRectItemgubun = itemgubun
ioffitem.FRectItemId = itemid
ioffitem.FRectItemOption = itemoption
ioffitem.GetOffOneItem

dim IsOnlineItem
	IsOnlineItem = (itemgubun="10")

dim opartner
set opartner = new CPartnerUser
if (ioffitem.FResultCount>0) then
    opartner.FRectDesignerID = ioffitem.FOneItem.Fmakerid
    opartner.GetOnePartnerNUser
end if

dim ooffontract
set ooffontract = new COffContractInfo
if (ioffitem.FResultCount>0) then
    ooffontract.FRectDesignerID = ioffitem.FOneItem.Fmakerid
    ooffontract.GetPartnerOffContractInfo
end if
%>
<script type='text/javascript'>

function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function EditItem(frm){
<% if (itemgubun<>"00") then %>
	if (frm.cd1.value.length<1){
		alert('카테고리를 선택하세요.');
		return;
	}
<% end if %>
	if (frm.shopitemname.value.length<1){
		alert('상품명을 입력하세요.');
		frm.shopitemname.focus();
		return;
	}

    if (frm.orgsellprice.value.length<1){
		alert('소비자가를 입력하세요.');
		frm.orgsellprice.focus();
		return;
	}

	if (frm.shopitemprice.value.length<1){
		alert('판매가를 입력하세요.');
		frm.shopitemprice.focus();
		return;
	}

	if (frm.shopsuplycash.value.length<1){
		alert('매입가를 입력하세요.');
		frm.shopsuplycash.focus();
		return;
	}

<% if (itemgubun="60") then %>
    if (frm.orgsellprice.value.substr(0,1) != '-'){
		frm.orgsellprice.value = "-"+frm.orgsellprice.value
	}
    if (frm.shopitemprice.value.substr(0,1) != '-'){
		frm.shopitemprice.value = "-"+frm.shopitemprice.value
	}
<% elseif (itemgubun="80") then %>
    if (frm.shopitemprice.value > 0){
		alert("사은품은 판매가가 0이하여야 합니다.");
		frm.shopitemprice.focus();
		return;
	}
    if (frm.orgsellprice.value > 0){
		alert("사은품은 소비자가 0이하여야 합니다.");
		frm.orgsellprice.focus();
		return;
	}
	if (frm.shopitemname.value.match(/^\[사은품\] /) == null) {
		alert("사은품 문구는 삭제할 수 없습니다.");
		return;
	}
<% elseif (itemgubun<>"00") then %>
    if (!IsDigit(frm.shopitemprice.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.shopitemprice.focus();
		return;
	}

    if (!IsDigit(frm.orgsellprice.value)){
		alert('소비자가는 숫자만 가능합니다.');
		frm.orgsellprice.focus();
		return;
	}

<% else %>
	if (!IsInteger(frm.shopitemprice.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.shopitemprice.focus();
		return;
	}

    if (!IsInteger(frm.orgsellprice.value)){
		alert('소비자가는 숫자만 가능합니다.');
		frm.orgsellprice.focus();
		return;
	}
<% end if %>

<% if (itemgubun<>"80") then %>
    if (frm.orgsellprice.value*1<frm.shopitemprice.value*1){
        alert('소비자가보다 실 판매가가 클 수 없습니다. 다시 입력하세요.');
		frm.shopitemprice.focus();
		return;
    }
<% end if %>

    if ((!frm.centermwdiv[0].checked)&&(!frm.centermwdiv[1].checked)){
        alert('센터 매입 구분을 선택 하세요.');
		frm.centermwdiv[0].focus();
		return;
    }

    if ((!frm.vatinclude[0].checked)&&(!frm.vatinclude[1].checked)){
        alert('과세 구분을 선택 하세요.');
		frm.vatinclude[0].focus();
		return;
    }

<% if Not IsOnlineItem then %>
//	if (frm.ioffimgmain.fileSize<1){
//		alert('이미지를 입력해 주세요 - 필수 사항입니다.');
//		frm.file1.focus();
//		return;
//	}
<% end if %>
	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}
}

function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640 height=540");
	popwin.focus();
}

// ============================================================================
// 카테고리등록
function editCategory(cdl,cdm,cdn){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cdn=" + cdn ;

	popwin = window.open('/common/module/categoryselect.asp?' + param ,'editcategory','width=700,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function setCategory(cd1,cd2,cd3,cd1_name,cd2_name,cd3_name){
	var frm = document.frmedit;
	frm.cd1.value = cd1;
	frm.cd2.value = cd2;
	frm.cd3.value = cd3;
	frm.cd1_name.value = cd1_name;
	frm.cd2_name.value = cd2_name;
	frm.cd3_name.value = cd3_name;
}


</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		상품코드 : <input type="text" class="text" name="barcode" value="<%= barcode %>" size="20">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
<% if (ioffitem.FResultCount<1) then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center">[검색 결과가 없습니다.]</td>
</tr>
<% else %>
<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>
<form name="frmedit" method=post action="offitemedit_process.asp" >
<input type=hidden name=itemgubun value="<%= itemgubun %>">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoption value="<%= itemoption %>">

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">상품코드</td>
	<td bgcolor="#FFFFFF" colspan="2"><%= ioffitem.FOneItem.GetBarcode %>

	<%if left(ioffitem.FOneItem.GetBarcode,2) = "10" then %>
		온라인공용상품
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "90" then %>
		오프라인전용상품
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "95" then %>
		가맹점개별매입판매상품
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "80" then %>
		사은품
	<% elseif left(ioffitem.FOneItem.GetBarcode,2) = "70" then %>
		소모품
	<% end if %>
	<br><font color="#AAAAAA">(90오프라인전용, 80사은품, 70소모품, 95가맹점개별매입판매)</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>상품명</td>
	<td bgcolor="#FFFFFF" colspan="2">
	<input type="text" class="text" name="shopitemname" value="<%= ioffitem.FOneItem.Fshopitemname %>" size="40" maxlength="40">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>옵션명</td>
	<% if (IsOnlineItem) and (ioffitem.FOneItem.Fitemoption<>"0000") then %>
	<td bgcolor="#FFFFFF" colspan="2">
		<input type="text" class="text" name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>" size="20" maxlength="20" class="input_01" >
	</td>
	<% else %>
		<td bgcolor="#FFFFFF" colspan="2">
			<input type="text" name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>" size="40" maxlength="40" class="input_01">
		</td>
	<% end if %>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td>카테고리</td>
	<td bgcolor="#FFFFFF" colspan="2">
	  <input type="hidden" name="cd1" value="<%= ioffitem.FOneItem.FCateCDL %>">
	  <input type="hidden" name="cd2" value="<%= ioffitem.FOneItem.FCateCDM %>">
	  <input type="hidden" name="cd3" value="<%= ioffitem.FOneItem.FCateCDS %>">

      <input type="text" class="text" name="cd1_name" value="<%= ioffitem.FOneItem.FCateCDLName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd2_name" value="<%= ioffitem.FOneItem.FCateCDMName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" class="text" name="cd3_name" value="<%= ioffitem.FOneItem.FCateCDSName %>" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" class="button" value="선택" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>가격설정</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<td bgcolor="#FFFFFF" >소비자가</td>
				<td bgcolor="#FFFFFF" >실 판매가</td>
				<% if not(C_IS_SHOP) then %>
					<td bgcolor="#FFFFFF" >매입가</td>
				<% end if %>
				<td bgcolor="#FFFFFF" >공급가</td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
			    <td bgcolor="#FFFFFF"><input type=text name="orgsellprice" value="<%= ioffitem.FOneItem.FShopItemOrgprice %>" size=8 maxlength=9 class="input_right" ></td>
				<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="<%= ioffitem.FOneItem.Fshopitemprice %>" size=8 maxlength=9 class="input_right" ></td>
				<% if not(C_IS_SHOP) then %>
					<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="<%= ioffitem.FOneItem.Fshopsuplycash %>" size=8 maxlength=9 class="input_right" ></td>
				<% end if %>
				<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="<%= ioffitem.FOneItem.Fshopbuyprice %>" size=8 maxlength=9 class="input_right" ></td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
				<td bgcolor="#FFFFFF" colspan="2"></td>
				<td bgcolor="#FFFFFF" colspan="2">(0 인경우 기본마진 자동 설정)</td>
			</tr>
			<tr bgcolor="#DDDDFF" align="center">
			    <td bgcolor="#FFFFFF" colspan="4">
			        <% if (ioffitem.FOneItem.FItemGubun="10") then %>
			            <b>온라인 판매 상품의 경우 익일 새벽에 온라인 판매가와<br>동일하게 설정됩니다.</b>
			        <% end if %>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>사용유무</td>
	<td bgcolor="#FFFFFF">
		<% if ioffitem.FOneItem.Fisusing="Y" then %>
		<input type=radio name=isusing value="Y" checked >사용함
		<input type=radio name=isusing value="N">사용안함
		<% else %>
		<input type=radio name=isusing value="Y"  >사용함
		<input type=radio name=isusing value="N" checked >사용안함
		<% end if %>
	</td>
	<td rowspan="4" bgcolor="#FFFFFF" align="center">
		<% if IsOnlineItem then %>
		<img src="<%= ioffitem.FOneItem.FimageList %>" width="100" height="100">
		<% else %>
		<a href="javascript:popOffImageEdit('<%= ioffitem.FOneItem.GetBarcode %>');"><img src="<%= ioffitem.FOneItem.FOffImgList %>" width="100" height="100" border="0"></a>
        <br>
        <a href="javascript:popOffImageEdit('<%= ioffitem.FOneItem.GetBarcode %>');">[이미지수정]</a>
		<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>센터매입구분</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="centermwdiv" value="W" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="W","checked","") %> >위탁
		<input type="radio" name="centermwdiv" value="M" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="M","checked","") %> >매입
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>과세구분</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="vatinclude" value="Y" <%= ChkIIF(ioffitem.FOneItem.Fvatinclude="Y","checked","") %>  >과세
	<input type="radio" name="vatinclude" value="N" <%= ChkIIF(ioffitem.FOneItem.Fvatinclude="N","checked","") %> > <font color="<%= ChkIIF(ioffitem.FOneItem.Fvatinclude="N","blue","#000000") %>">면세</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td>범용바코드</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="extbarcode" value="<%= ioffitem.FOneItem.Fextbarcode %>" size="20" maxlength="20" class="input_01" >
	</td>
</tr>

<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>

<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>브랜드계약정보</td>
	<td bgcolor="#FFFFFF" colspan="2"><a href="javascript:PopUpcheInfo('<%= ioffitem.FOneItem.Fmakerid %>');"><%= ioffitem.FOneItem.Fmakerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	<br><font color="#AAAAAA">(브랜드 변경시 관리자에게 문의)</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td>온라인</td>
	<td bgcolor="#FFFFFF" colspan="2">
		<%= FormatNumber(ioffitem.FOneItem.FOnlineOrgprice,0) %> / <%= FormatNumber(ioffitem.FOneItem.FOnlineBuycash,0) %>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(ioffitem.FOneItem.FmwDiv) %>"><%= mwdivName(ioffitem.FOneItem.FmwDiv) %></font>
		&nbsp;
		<% if ioffitem.FOneItem.FOnlineSellcash<>0 then %>
		<%= CLng((1- ioffitem.FOneItem.FOnlineBuycash/ioffitem.FOneItem.FOnlineOrgprice)*100) %> %
		<% end if %>

		<% if (ioffitem.FOneItem.FOnlineSailYn="Y") then %>
		<br>
		<font color="red">
		<%= FormatNumber(ioffitem.FOneItem.FOnlineSellcash,0) %> / <%= FormatNumber(ioffitem.FOneItem.FOnlineBuycash,0) %>
		&nbsp;&nbsp;
			<% if (ioffitem.FOneItem.FOnlineOrgprice<>0) then %>
		        <%= CLng((ioffitem.FOneItem.FOnlineOrgprice - ioffitem.FOneItem.FOnlineSellcash)/ioffitem.FOneItem.FOnlineOrgprice*100) %>%
		    <% end if %>
		    할인
		</font>
		&nbsp;&nbsp;
		<font color="<%= mwdivColor(ioffitem.FOneItem.FmwDiv) %>"><%= mwdivName(ioffitem.FOneItem.FmwDiv) %></font>
		&nbsp;
			<% if ioffitem.FOneItem.FOnlineSellcash<>0 then %>
				<%= CLng((1- ioffitem.FOneItem.FOnlineBuycash/ioffitem.FOneItem.FOnlineSellcash)*100) %> %
			<% end if %>

		<% end if %>

	</td>
</tr>


<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>오프라인<br>[직영점]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= ioffitem.FOneItem.Fmakerid %>')"><b>직영점대표</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td width=60><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td width=60><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>오프라인<br>[가맹점]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop800','<%= ioffitem.FOneItem.Fmakerid %>')"><b>가맹점점대표</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>

		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("pink") %>">
	<td width=100>오프라인<br>[해외공급]</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop870','<%= ioffitem.FOneItem.Fmakerid %>')"><b>해외공급대표</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop870") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop870") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="5")  then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= ioffitem.FOneItem.Fmakerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>

<tr height="1" bgcolor="FFFFFF">
	<td colspan="15"></td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>등록일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>최종수정일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fupdt %></td>
</tr>

</form>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align=center>
		<% if not(C_IS_SHOP) then %>
			<input type="button" class="button" value=" 저장 " onclick="EditItem(frmedit)">
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->