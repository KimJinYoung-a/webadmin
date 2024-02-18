<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim itemgubun,itemid, itemoption, barcode
barcode	  = RequestCheckVar(request("barcode"),20)

itemgubun = Left(barcode,2)
itemid	  = CLng(Mid(barcode,3,6))
itemoption = Right(barcode,4)
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
opartner.FRectDesignerID = ioffitem.FOneItem.Fmakerid
opartner.GetOnePartnerNUser

dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = ioffitem.FOneItem.Fmakerid
ooffontract.GetPartnerOffContractInfo

dim i
%>
<script language='javascript'>
function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function EditItem(frm){
//alert('잠시 수정중입니다.');
//return;
	if (frm.cd1.value.length<1){
		alert('카테고리를 선택하세요.');
		return;
	}

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
	
    if (frm.orgsellprice.value*1<frm.shopitemprice.value*1){
        alert('소비자가보다 실 판매가가 클 수 없습니다. 다시 입력하세요.');
		frm.shopitemprice.focus();
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
function editCategory(cdl,cdm,cds){
	var param = "cdl=" + cdl + "&cdm=" + cdm + "&cds=" + cds ;

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
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#FFFFFF>
<tr>
	<td>&gt;&gt;오프라인 상품 수정</td>
</tr>
</table>

<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<form name="frmedit" method=post action="offitemedit_process.asp" >
<input type=hidden name=itemgubun value="<%= itemgubun %>">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoption value="<%= itemoption %>">

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">상품코드</td>
	<td bgcolor="#FFFFFF" colspan="4"><%= ioffitem.FOneItem.GetBarcode %>
	
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
	<br><font color="#AAAAAA">(90오프라인전용, 80이벤트 ,70소모품, 95가맹점개별매입판매)</font>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>상품명</td>
	<td bgcolor="#FFFFFF" colspan=4>
	<input type=text name="shopitemname" value="<%= ioffitem.FOneItem.Fshopitemname %>" size=40 maxlength=30 class="input_01" >
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>옵션명</td>
	<% if (IsOnlineItem) and (ioffitem.FOneItem.Fitemoption<>"0000") then %>
	<td bgcolor="#FFFFFF" colspan=4><input type=text name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>" size=20 maxlength=20 class="input_01" ></td>
	<% else %>
	<input type=hidden name="shopitemoptionname" value="<%= ioffitem.FOneItem.Fshopitemoptionname %>">
	<td bgcolor="#FFFFFF" colspan=4><%= ioffitem.FOneItem.Fshopitemoptionname %></td>
	<% end if %>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td width=100 >카테고리</td>
	<td bgcolor="#FFFFFF" colspan=4>
	  <input type="hidden" name="cd1" value="<%= ioffitem.FOneItem.FCateCDL %>">
	  <input type="hidden" name="cd2" value="<%= ioffitem.FOneItem.FCateCDM %>">
	  <input type="hidden" name="cd3" value="<%= ioffitem.FOneItem.FCateCDS %>">

      <input type="text" name="cd1_name" class="text" value="<%= ioffitem.FOneItem.FCateCDLName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" class="text" value="<%= ioffitem.FOneItem.FCateCDMName %>" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" class="text" value="<%= ioffitem.FOneItem.FCateCDSName %>" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" class="button" value="선택" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td width=100 align="left" rowspan="4">가격설정</td>
	<td bgcolor="#FFFFFF" >소비자가</td>
	<td bgcolor="#FFFFFF" >실 판매가</td>
	<td bgcolor="#FFFFFF" colspan="2">매입가</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
    <td bgcolor="#FFFFFF"><input type=text name="orgsellprice" value="<%= ioffitem.FOneItem.FShopItemOrgprice %>" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly></td>
	<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="<%= ioffitem.FOneItem.Fshopitemprice %>" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly></td>
	<td bgcolor="#FFFFFF" colspan="2"><input type=text name="shopsuplycash" value="<%= ioffitem.FOneItem.Fshopsuplycash %>" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly></td>
    <input type="hidden" name="shopbuyprice" value="<%= ioffitem.FOneItem.Fshopbuyprice %>">
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


<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>사용유무</td>
	<td bgcolor="#FFFFFF" colspan=3>
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
	<td width=100>센터매입구분</td>
	<td bgcolor="#FFFFFF" colspan=3>
	    <input type=radio name=centermwdiv value="W" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="W","checked","") %> >특정
	    <input type=radio name=centermwdiv value="M" <%= ChkIIF(ioffitem.FOneItem.FCenterMwDiv="M","checked","") %>>매입
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" >
	<td width=100 >과세구분</td>
	<td bgcolor="#FFFFFF" colspan=3>
	<% if ioffitem.FOneItem.Fvatinclude="Y" then %>
	<input type=radio name=vatinclude value="Y" checked >과세
	<input type=radio name=vatinclude value="N">면세
	<% else %>
	<input type=radio name=vatinclude value="Y"  >과세
	<input type=radio name=vatinclude value="N" checked >면세
	<% end if %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width=100>범용바코드</td>
	<td bgcolor="#FFFFFF" colspan=3><input type=text name="extbarcode" value="<%= ioffitem.FOneItem.Fextbarcode %>" size=20 maxlength=20 class="input_01" ></td>
</tr>
</table>
<p>
<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>

<tr bgcolor="#FFDDDD">
	<td width=100>브랜드계약정보</td>
	<td bgcolor="#FFFFFF" colspan=5><a href="javascript:PopUpcheInfo('<%= ioffitem.FOneItem.Fmakerid %>');"><%= ioffitem.FOneItem.Fmakerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<% if IsOnlineItem then %>
<tr bgcolor="#FFDDDD">
	<td width=100 rowspan=3>온라인</td>
	<td bgcolor="#FFFFFF" colspan=5><%= opartner.FOneItem.GetMWUName %> &nbsp;&nbsp; <%= opartner.FOneItem.Fdefaultmargine %> %</td>
</tr>

<tr bgcolor="#FFDDDD">
	<td bgcolor="#FFFFFF" >소비자가</td>
	<td bgcolor="#FFFFFF" >판매가</td>
	<td bgcolor="#FFFFFF" >매입가</td>
	<td bgcolor="#FFFFFF" >마진</td>
</tr>

<tr bgcolor="#FFDDDD">
	<td bgcolor="#FFFFFF" align=right><%= ioffitem.FOneItem.FOnlineOrgprice %></td>
	<td bgcolor="#FFFFFF" align=right>
	<% if (ioffitem.FOneItem.FOnlineSailYn="Y") then %>
	<font color=red><%= ioffitem.FOneItem.FOnlineSellcash %></font>
	<% else %>
	<%= ioffitem.FOneItem.FOnlineSellcash %>
	<% end if %>
	</td>
	<td bgcolor="#FFFFFF" align=right><%= ioffitem.FOneItem.FOnlineBuycash %></td>
	<td bgcolor="#FFFFFF" align=center>
	<% if ioffitem.FOneItem.FOnlineSellcash<>0 then %>
	<%= 100-CLng(ioffitem.FOneItem.FOnlineBuycash/ioffitem.FOneItem.FOnlineSellcash*100*100)/100 %> %
	<% end if %>
	</td>
</tr>
<% end if %>

<tr bgcolor="#FFDDDD">
	<td width=100>오프라인-직영</td>
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
<tr bgcolor="#FFDDDD">
	<td width=100>오프라인-가맹</td>
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
<tr bgcolor="#DDDDFF">
	<td width=100>등록일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fregdate %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>최종수정일</td>
	<td bgcolor="#FFFFFF" colspan=5><%= ioffitem.FOneItem.Fupdt %></td>
</tr>
</form>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center><input type=button value=" 저  장 " onclick="EditItem(frmedit)" class="input_01"></td>
</tr>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
set ioffitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->