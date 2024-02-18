<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  상품 쿠폰
' History : 2010.09.30 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/itemcouponcls.asp" -->
<%
dim itemcouponidx ,oitemcouponmaster ,IsEditMode, IsExpiredCoupon
	itemcouponidx = requestCheckVar(request("itemcouponidx"),9)
	if itemcouponidx="" then itemcouponidx=0

set oitemcouponmaster = new CItemCouponMaster
	oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
	oitemcouponmaster.GetOneItemCouponMaster()

IsEditMode = (CStr(itemcouponidx)<>"0")
%>

<script language='javascript'>

function OpenCouponMaster(){
	frmcoupon.mode.value="opencoupon";

	if (confirm('쿠폰을 오픈 하시겠습니까?')){
		frmcoupon.submit();
	}
}

function reserveCouponMaster(){
	frmcoupon.mode.value="reservecoupon";

	if (confirm('쿠폰오픈을 예약 하시겠습니까?')){
		frmcoupon.submit();
	}

}

var alertCnt = 0;
function AlertMarginChange(){
	if (alertCnt==0){
		alert('마진 설정을 변경하시면 대상상품 전체에 적용 됩니다.');
		alertCnt++;
	}
}

function CloseCouponMaster(){
	frmcoupon.mode.value="closecoupon";

	if (confirm('!! 고객께 발행된 쿠폰도 같이 종료 됩니다.\n\n쿠폰을 강제 종료 하시겠습니까?')){
		frmcoupon.submit();
	}
}

function fninput(v){

	var ele = document.getElementById('marginlayer');

	if (v==20){
		ele.style.display="";
	}else {
		ele.style.display="none";
	}
}

function SaveCouponMaster(frm, isEditMode){
	if (frmcoupon.itemcouponname.value.length<2){
		alert('쿠폰명을 입력해 주세요.');
		frmcoupon.itemcouponname.focus();
		return;
	}

    if ((!frmcoupon.couponGubun[0].checked)&&(!frmcoupon.couponGubun[1].checked)&&(!frmcoupon.couponGubun[2].checked)){
        alert('쿠폰 구분을 선택하세요..');
		frmcoupon.couponGubun[0].focus();
		return;
    }

    if (frmcoupon.couponGubun[2].checked){
        alert('지정 쿠폰 발행시 시스템팀  문의 요망!');
    }

	if (frmcoupon.itemcouponvalue.value.length<1){
		alert('쿠폰 금액 또는 할인율을 입력해 주세요.');
		frmcoupon.itemcouponvalue.focus();
		return;
	}

	if (!IsDigit(frmcoupon.itemcouponvalue.value)){
		alert('쿠폰 금액 또는 할인율은 숫자만 가능합니다.');
		frmcoupon.itemcouponvalue.focus();
		return;
	}


	if ((!frmcoupon.itemcoupontype[0].checked)&&(!frmcoupon.itemcoupontype[1].checked)&&(!frmcoupon.itemcoupontype[2].checked)){
		alert('할인 타입을 설정해 주세요.');
		frmcoupon.itemcouponvalue.focus();
		return;
	}

    if ((frmcoupon.itemcoupontype[2].checked)&&(frmcoupon.itemcouponvalue.value!='2000')){
		alert('무료배송 쿠폰은 할인액은 2000원 입니다.');
		frmcoupon.itemcouponvalue.focus();
		return;
	}

	if ((frmcoupon.itemcoupontype[2].checked)&&!(frmcoupon.margintype.value=='20'||frmcoupon.margintype.value=='50'||frmcoupon.margintype.value=='80')){
		alert('무료배송 쿠폰 발급시 반반부담, 직접설정 또는 무료배송500업체부담으로 선택해주세요.');
		frmcoupon.margintype.focus();
		return;
	}

	if (frmcoupon.itemcouponstartdate.value.length!=10){
		alert('쿠폰 발급 시작일을 입력해 주세요.');
		frmcoupon.itemcouponstartdate.focus();
		return;
	}

	if (frmcoupon.itemcouponstartdate2.value.length!=8){
		alert('쿠폰 발급 시작일을 입력해 주세요.');
		frmcoupon.itemcouponstartdate2.focus();
		return;
	}

	if (frmcoupon.itemcouponexpiredate.value.length!=10){
		alert('쿠폰 발급 종료일을 입력해 주세요.');
		frmcoupon.itemcouponexpiredate.focus();
		return;
	}

	if (frmcoupon.itemcouponexpiredate2.value.length!=8){
		alert('쿠폰 발급 종료일을 입력해 주세요.');
		frmcoupon.itemcouponexpiredate2.focus();
		return;
	}

	if (frmcoupon.margintype.value.length<1){
		alert('마진 구분을 설정해 주세요.');
		frmcoupon.margintype.focus();
		return;
	}

	if (frmcoupon.margintype.value==20){
		if (frmcoupon.defaultmargin.value.length<1){
			alert('마진을 입력해 주세요.');
			frmcoupon.defaultmargin.focus();
			return;
		}

	}

    if (isEditMode){
        if (confirm('수정 하시겠습니까?')){
    		frmcoupon.submit();
    	}
    }else{
    	if (confirm('저장 하시겠습니까?')){
    		frmcoupon.submit();
    	}
    }
}

</script>

<!---- 검색 ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		쿠폰번호 : <input type="text" name="itemcouponidx" value="<%= itemcouponidx %>" Maxlength="12" size="12" readonly >
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>	
</form>
</table>
<!---- /검색 ---->

<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#BABABA">
<form name="frmcoupon" method="post" action="itemcoupon_Process.asp">
<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>">
<input type="hidden" name="mode" value="couponmaster">
<tr bgcolor="#DDDDFF">
	<td width="100">쿠폰명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="itemcouponname" value="<%= oitemcouponmaster.FOneItem.Fitemcouponname %>" size="40" maxlength="30"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100">쿠폰구분</td>
	<td bgcolor="#FFFFFF">
	    <input type="radio" name="couponGubun" value="C" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="C","checked","") %> >일반
	    <input type="radio" name="couponGubun" value="T" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="T","checked","") %> >타겟(E-mail특가)
	    <input type="radio" name="couponGubun" value="P" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="P","checked","") %> >지정인발급(프런트 발급 불가 : 시스템팀 문의)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >할인율</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="itemcouponvalue" value="<%= oitemcouponmaster.FOneItem.Fitemcouponvalue %>" size="6">
		<input type="radio" name="itemcoupontype" value="1" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="1" then response.write "checked" %> > %
		<input type="radio" name="itemcoupontype" value="2" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="2" then response.write "checked" %> > 원
		<input type="radio" name="itemcoupontype" value="3" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="3" then response.write "checked" %> > 배송료할인쿠폰 (2000 입력)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >적용기간</td>
	<td bgcolor="#FFFFFF">
	<input type="text" class="text" name="itemcouponstartdate" value="<%= Left(oitemcouponmaster.FOneItem.Fitemcouponstartdate,10) %>" size="10" maxlength="10">
	<input type="text" class="text_ro" name="itemcouponstartdate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Fitemcouponstartdate<>"",Right(oitemcouponmaster.FOneItem.Fitemcouponstartdate,8),"00:00:00") %>" size="8" maxlength="8">
	<a href="javascript:calendarOpen(frmcoupon.itemcouponstartdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	~
	<input type="text" class="text" name="itemcouponexpiredate" value="<%= Left(oitemcouponmaster.FOneItem.Fitemcouponexpiredate,10) %>" size="10" maxlength="10">
	<input type="text" class="text_ro" name="itemcouponexpiredate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Fitemcouponexpiredate<>"",Right(oitemcouponmaster.FOneItem.Fitemcouponexpiredate,8),"23:59:59") %>" size="8" maxlength="8">
	<a href="javascript:calendarOpen(frmcoupon.itemcouponexpiredate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	<br>(<%= Left(now(),10) %> 00:00:00)  ~  (<%= Left(now(),10) %> 23:59:59)
	<br><font color="#808080">(※ 고객이 이미 다운로드한 쿠폰은 적용기간이 변경되지 않습니다. 따라서 기간 수정시에 유의해주세요.)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >기본 마진구분</td>
	<td bgcolor="#FFFFFF">
		<select name="margintype" onchange="AlertMarginChange();fninput(this.value);">
		<!--<option value="">---선택--- -->
		<option value="30" <% if oitemcouponmaster.FOneItem.Fmargintype="30" then response.write "selected" %> >동일마진
		<option value="60" <% if oitemcouponmaster.FOneItem.Fmargintype="60" then response.write "selected" %> >업체부담
		<option value="50" <% if oitemcouponmaster.FOneItem.Fmargintype="50" then response.write "selected" %> >반반부담
		<option value="10" <% if oitemcouponmaster.FOneItem.Fmargintype="10" then response.write "selected" %> >텐바이텐부담
		<option value="20" <% if oitemcouponmaster.FOneItem.Fmargintype="20" then response.write "selected" %> >직접설정
		<option value="00" <% if oitemcouponmaster.FOneItem.Fmargintype="00" then response.write "selected" %> >상품개별설정
		<option value="90" <% if oitemcouponmaster.FOneItem.Fmargintype="90" then response.write "selected" %> >20%전체행사
		<option value="80" <% if oitemcouponmaster.FOneItem.Fmargintype="80" then response.write "selected" %> >무료배송(500업체부담)
		</select>
		<span id="marginlayer" style="display:<% IF oitemcouponmaster.FOneItem.Fmargintype<>"20" Then response.write "none" %>"><input type="text" class="text" name="defaultmargin" value="<%=oitemcouponmaster.FOneItem.FDefaultMargin%>" size="3" maxlength="3" onChange="AlertMarginChange();">%</span>
		<font color="#808080">(상품별로 마진율이 다른 경우 별도로 설정 가능합니다.)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >쿠폰설명</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="itemcouponexplain" value="<%= oitemcouponmaster.FOneItem.Fitemcouponexplain %>" size="60" maxlength="50">
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >발급 상태</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName %>
	<% if (oitemcouponmaster.FOneItem.FItemCouponIdx>0) then %>
    	<% if (oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
    		--&gt;<input type="button" value="오픈" onclick="OpenCouponMaster();" class="button" >
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="0")  then %>
    		--&gt;<input type="button" value="발급예약" onclick="reserveCouponMaster();" class="button">
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="9")  then %>

    	<% else %>
    	--&gt;<input type="button" value="발급강제종료" onclick="CloseCouponMaster();" class="button"  >
    	(종료일 12시 15분에 자동 종료됩니다.)
    	<% end if %>
    <% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >등록일</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.Fregdate %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<% if (IsEditMode) then %>
	    <% if (oitemcouponmaster.FOneItem.Fopenstate="0") then %>
	    <td colspan="2" align="center"><input type="button" value="수 정" onclick="SaveCouponMaster(frmcoupon, true)" class="button"></td>
	    <% elseif (Not oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
	    <td colspan="2" align="center"><input type="button" value="수 정" onclick="SaveCouponMaster(frmcoupon, true)" class="button" Disabled ></td>
	    <% else %>
	    <td colspan="2" align="center"><input type="button" value="수 정" onclick="SaveCouponMaster(frmcoupon, true)" class="button"></td>
	    <% end if %>
	<% else %>
	<td colspan="2" align="center"><input type="button" value="저 장" onclick="SaveCouponMaster(frmcoupon, false)" class="button"></td>
	<% end if %>
</tr>
</form>
</table>

<%
	set oitemcouponmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->