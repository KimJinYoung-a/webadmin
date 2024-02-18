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
dim makerid
makerid = session("ssBctID")


dim opartner
set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if makerid<>"" then
	opartner.GetOnePartnerNUser
end if

dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = makerid

if makerid<>"" then
	ooffontract.GetPartnerOffContractInfo
end if

dim i

''DefaultCenterMwdiv  
dim DefaultCenterMwdiv
DefaultCenterMwdiv = GetDefaultItemMwdivByBrand(makerid)
%>
<script language='javascript'>
function AttachImage(comp,filecomp){
	comp.src=filecomp.value;
}

function ChangeBrand(comp){
	location.href="?makerid=" + comp.value;
}

function CheckAddItem(frm){
	if ((frm.itemgubun[0].checked==false)&&(frm.itemgubun[1].checked==false)){
		alert('상품구분을 선택하세요.');
		return;
	}

	if (frm.cd1.value.length<1){
		alert('카테고리를 선택하세요.');
		return;
	}

	if (frm.makerid.value.length<1){
		alert('브랜드를 선택하세요.');
		return;
	}

	if (frm.shopitemname.value.length<1){
		alert('상품명을 입력하세요.');
		frm.shopitemname.focus();
		return;
	}

	if ((frm.extbarcode.value.length>0) && (frm.extbarcode.value.length<10)){
		alert('바코드 길이가 너무 짧습니다.  범용 바코드가 있는경우만 입력해 주세요');
		frm.extbarcode.focus();
		return;
	}

	if (!IsDigit(frm.shopitemprice.value)){
		alert('판매가는 숫자만 가능합니다.');
		frm.shopitemprice.focus();
		return;
	}


//	if (!IsDigit(frm.discountsellprice.value)){
//		alert('할인 판매가는 숫자만 가능합니다.');
//		frm.discountsellprice.focus();
//		return;
//	}


	if (!IsDigit(frm.shopsuplycash.value)){
		alert('업체 매입가는 숫자만 가능합니다.');
		frm.shopsuplycash.focus();
		return;
	}

	if (!IsDigit(frm.shopbuyprice.value)){
		alert('샾 공급가는 숫자만 가능합니다.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.shopsuplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! 기본 계약 마진과 다를 경우에만 매입가 공급가를 입력 하셔야 합니다. \n\n계속 하시겠습니까?')){
			return;
		}
	}

    if (frm.file1.value.length<1){
		alert('이미지를 입력해 주세요 - 필수 사항입니다.');
		frm.file1.focus();
		return;
	}
    
//	if (frm.ioffimgmain.src.length<1){
//		alert('이미지를 입력해 주세요 - 필수 사항입니다.');
//		frm.file1.focus();
//		return;
//	}

//	if (!((/(.jpg|.jpeg)$/i).test(frm.ioffimgmain.src))){
//	  alert('기본 이미지는 jpg 파일만 가능합니다.');
//	}

	var ret = confirm('오프라인 전용 상품으로 등록 하시겠습니까?');

	if (ret) {
		frm.submit();
	}
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
	<td>&gt;&gt;오프라인 상품 등록</td>
</tr>
</table>

<table border=0 cellspacing=1 cellpadding=2 width=460 class="a" bgcolor=#3d3d3d>
<% if application("Svr_Info")="Dev" then %>	
<form name="frmedit" method=post action="http://testpartner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% else %>
<form name="frmedit" method=post action="http://partner.10x10.co.kr/linkweb/dooffitemimageeditwithdata.asp" enctype="MULTIPART/FORM-DATA">
<% end if %>
<input type=hidden name=mode value="addnewoffitem">
<input type=hidden name=makerid value="<%= makerid %>">

<tr bgcolor="#FFDDDD">
	<td width=100>브랜드계약정보</td>
	<td bgcolor="#FFFFFF" colspan=5><a href="javascript:PopUpcheInfo('<%= makerid %>');"><%= makerid %></a> (<%= opartner.FOneItem.Fsocname_kor %>,<%= opartner.FOneItem.FCompany_name %>)
	</td>
</tr>
<tr bgcolor="#FFDDDD">
	<td width=100 >온라인</td>
	<td bgcolor="#FFFFFF" colspan=5><%= opartner.FOneItem.GetMWUName %> &nbsp;&nbsp; <%= opartner.FOneItem.Fdefaultmargine %> %</td>
</tr>

<tr bgcolor="#FFDDDD">
	<td width=100>오프라인-직영</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=80%>
		<tr>
			<td ><a href="javascript:editOffDesinger('streetshop000','<%= makerid %>')"><b>직영점대표</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="1") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
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
			<td ><a href="javascript:editOffDesinger('streetshop800','<%= makerid %>')"><b>가맹점점대표</b></a></td>
			<td width=60><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width=60><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="5") then %>
		<tr>
			<td ><a href="javascript:editOffDesinger('<%= ooffontract.FItemList(i).Fshopid %>','<%= makerid %>')"><%= ooffontract.FItemList(i).Fshopname %></a></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>상품구분</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type="radio" name="itemgubun" value="90" checked >오프샾 전용상품(90) &nbsp;
	<input type="radio" name="itemgubun" value="70">소모품(70)
	<br><font color="#AAAAAA">(90오프라인전용, 80이벤트 ,70소모품, 95가맹점개별매입판매)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >카테고리</td>
	<td bgcolor="#FFFFFF" colspan=5>
	  <input type="hidden" name="cd1" value="">
	  <input type="hidden" name="cd2" value="">
	  <input type="hidden" name="cd3" value="">

      <input type="text" name="cd1_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd2_name" value="" size="12" readonly style="background-color:#E6E6E6">
      <input type="text" name="cd3_name" value="" size="12" readonly style="background-color:#E6E6E6">

      <input type="button" value="선택" onclick="editCategory(frmedit.cd1.value,frmedit.cd2.value,frmedit.cd3.value);">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>상품명</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=text name="shopitemname" value="" size=40 maxlength=30 class="input_01" >
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>옵션명</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=hidden name="shopitemoptionname" value="">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>범용바코드</td>
	<td bgcolor="#FFFFFF" colspan=5><input type=text name="extbarcode" value="" size=20 maxlength=20 class="input_01" >(있는 경우만 등록)</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>사용유무</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=isusing value="Y" checked >사용함
	<input type=radio name=isusing value="N">사용안함
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=100>센터매입구분</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=centermwdiv value="W" <%= ChkIIF(DefaultCenterMwdiv<>"M","checked","") %> >특정
	<input type=radio name=centermwdiv value="M" <%= ChkIIF(DefaultCenterMwdiv="M","checked","") %>>매입
	</td>
</tr>
<tr bgcolor="#DDDDFF" >
	<td width=100 >과세구분</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=radio name=vatinclude value="Y" checked >과세
	<input type=radio name=vatinclude value="N">면세
	</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td width=100 align="left" rowspan="3">가격설정</td>
	<td bgcolor="#FFFFFF" >판매가</td>
	<td bgcolor="#FFFFFF" >매입가</td>
	<td bgcolor="#FFFFFF" >공급가</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF"><input type=text name="shopitemprice" value="" size=8 maxlength=9 class="input_right" ></td>
	<td bgcolor="#FFFFFF"><input type=text name="shopsuplycash" value="0" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly ></td>
	<td bgcolor="#FFFFFF" ><input type=text name="shopbuyprice" value="0" size=8 maxlength=9 class="input_right" style="background-color : #DDDDDD" readonly ></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
	<td bgcolor="#FFFFFF" ></td>
	<td bgcolor="#FFFFFF" colspan="3">(0 인경우 기본마진 으로 자동 설정됨)</td>
</tr>

</tr>
<tr bgcolor="#DDDDFF">
	<td width=100 valign=top>오프상품<br>이미지</td>
	<td bgcolor="#FFFFFF" colspan=5>
	<input type=file name=file1 class="input_01" size=20 onchange="AttachImage(ioffimgmain,this)" >(400 x 400 px)
	<br>(기본 이미지는 꼭 <b>jpg</b> 파일로 올려주시기 바랍니다.)
	<img name="ioffimgmain" src="" width=340 height=340>
	</td>
</tr>
</form>
<tr bgcolor="#FFFFFF">
	<td colspan=6 align=center><input type=button value=" 저  장 " onclick="CheckAddItem(frmedit)" class="input_01"></td>
</tr>
</table>

<%
set opartner = Nothing
set ooffontract = Nothing
%>

<script language='javascript'>
alert('온라인에서 판매하거나 판매 예정인 상품은 절대 등록하지 마시기 바랍니다. \n 상품이 이중등록되어 재고관리에 어려움이 있습니다. \n 온라인상품 등록후 하루지나면 오프라인 상품으로 자동 등록됩니다. \n\n 꼭 필요한경우(오프라인에서만 판매하실 경우 등)만 등록해주세요');
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->