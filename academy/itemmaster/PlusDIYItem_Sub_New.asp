<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->
<%
dim itemid
itemid = requestCheckvar(request("itemid"),9)
itemid = CStr(itemid)
'itemid = Cint(itemid)

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if


dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FRectItemID = itemid

if itemid<>"" then
	oPlusSaleItem.GetOnePlusSaleSubItem
end if

dim i
dim IsNewReg        '' 신규등록인지
IsNewReg = (oPlusSaleItem.FResultCount<1)

'' 기존 IsLinkedItem 인경우
dim IsLinkedItem
if itemid<>"" then
    IsLinkedItem = oPlusSaleItem.IsPlusSaleLinkItem
end if
%>

<script language='javascript'>
function CalcuMargin(frm){
    var vSalePro = frm.PlusSalePro.value;
    var vMarginFlag = frm.PlusSaleMaginFlag.value;
    var vOrgMargin = 0;
    var vSaleMargin = 0;

    if (vSalePro.length<1){
        alert('할인율을 입력하세요.');
        frm.PlusSalePro.focus();
    }

    if (!IsDigit(vSalePro)){
        alert('할인율을 숫자로 입력하세요.');

        frm.PlusSalePro.focus();
        frm.PlusSalePro.select();
    }

    frm.tmpSellCash.value = parseInt(frm.osellcash.value-frm.osellcash.value*vSalePro/100);
    vOrgMargin = 100-parseInt(frm.obuycash.value*1/frm.osellcash.value*1*100*100)/100;

    frm.PlusSaleMargin.readOnly = true;
    frm.PlusSaleMargin.className = "text_ro";

    if (vMarginFlag=="1"){      //동일마진
        vSaleMargin = vOrgMargin;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);

    }else if(vMarginFlag=="2"){  //업체부담 : 사용안함
        frm.tmpBuyCash.value = frm.tmpSellCash.value*1-parseInt(frm.osellcash.value*1-frm.obuycash.value*1);
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-(frm.osellcash.value*1-frm.obuycash.value*1));

    }else if(vMarginFlag=="3"){  //반반부담 : 사용안함
        frm.tmpBuyCash.value = frm.obuycash.value*1-parseInt((frm.osellcash.value*1-frm.tmpSellCash.value*1)/2);
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);

    }else if(vMarginFlag=="4"){  //텐바이텐부담 : 실제 디비 마진 0 저장.(+-1언 오차 개연성?.)
        frm.tmpBuyCash.value = frm.obuycash.value*1;
        vSaleMargin = 100-parseInt(frm.tmpBuyCash.value*1/frm.tmpSellCash.value*1*100*100)/100;
        frm.PlusSaleMargin.value = vSaleMargin;
        //frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);
        frm.tmpBuyCash.value = frm.obuycash.value;

    }else if(vMarginFlag=="5"){  //직접설정
        frm.PlusSaleMargin.readOnly = false;
        frm.PlusSaleMargin.className = "text";
        frm.PlusSaleMargin.focus();

        //vSaleMargin = vOrgMargin;
        //frm.PlusSaleMargin.value = vSaleMargin;
        vSaleMargin = frm.PlusSaleMargin.value;
        frm.tmpBuyCash.value = Math.round(frm.tmpSellCash.value-frm.tmpSellCash.value*vSaleMargin/100);


    }



}

function setComp(comp){
    if (comp.name=="termsGubun"){
        if (comp.value=="A"){
            comp.form.PlusSaleStartDate.value = "1901-01-01";
            comp.form.PlusSaleEndDate.value = "9999-12-31";
        }else if (comp.value=="S"){
            comp.form.PlusSaleStartDate.value = "";
            comp.form.PlusSaleEndDate.value = "";
        }
    }
}

function RegPLusSale(frm){
    if (frm.PlusSalePro.value.length<1){
        alert('할인율을 입력하세요.');
        frm.PlusSalePro.focus();
        return;
    }

    if (!IsDigit(frm.PlusSalePro.value)){
        alert('할인율을 입력하세요.');
        frm.PlusSalePro.focus();
        return;
    }

    if (frm.PlusSalePro.value*1>50){
        alert('할인율을 확인해 주세요.');
        frm.PlusSalePro.focus();
        return;
    }

    if ((frm.omwdiv.value=="M")&&(frm.PlusSaleMaginFlag.value!="4")){
        alert('상품 매입구분이 매입인 경우 매입 마진 구분을 텐바이텐 부담으로 설정하세요.');
        frm.PlusSaleMaginFlag.focus();
        return;
    }

    if ((frm.PlusSaleMargin.value*1>100)||(frm.PlusSaleMargin.value*1<1)){
        alert('할인시 공급율을 확인해 주세요.');
        //frm.PlusSaleMargin.focus();
        return;
    }

    if  (!IsDouble(frm.PlusSaleMargin.value)){
        alert('할인시 공급율을 확인해 주세요.');
        //frm.PlusSaleMargin.focus();
        return;
    }

    if (frm.tmpBuyCash.value*1<0){
        alert('할인시 매입가를 확인해 주세요.');
        return;
    }

    if (frm.tmpSellCash.value*1<frm.tmpBuyCash.value*1){
        alert('할인시 매입가를 확인해 주세요. 할인시 판매가보다 클 수 없습니다.');
        return;
    }

    //PlusSaleMaginFlag

    if ((!frm.termsGubun[0].checked)&&(!frm.termsGubun[1].checked)){
        alert('기간 진행 여부를 선택해 주세요.');
        frm.termsGubun[1].focus();
        return;
    }


    if (frm.PlusSaleStartDate.value.length<1){
        alert('시작일을 선택 하세요.');
        return;
    }

    if (frm.PlusSaleEndDate.value.length<1){
        alert('종료일을 선택 하세요.');
        return;
    }

    if (frm.PlusSaleStartDate.value>frm.PlusSaleEndDate.value){
        alert('종료일을 시작일 이전으로 설정 할 수 없습니다.');
        return;
    }

    <% if IsNewReg then %>
    if (confirm('등록 하시겠습니까?')){
        frm.submit();
    }
    <% else %>
    if (confirm('수정 하시겠습니까?')){
        frm.submit();
    }
    <% end if %>
}

function showLinkedItemList(iitemid){
    var popwin = window.open('PlusDIYItem_Edit.asp?itemid=' + iitemid,'PlusDIYItem_Edit','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}


function DelPLusSale(frm){
    if (confirm('Plus Sale  추가구성 상품 을 삭제 하시겠습니까? - 메인상품 링크도 같이 삭제됩니다.')){
        frm.mode.value = "delPlusSale";
        frm.submit();
    }
}
</script>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method="get" >
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<img src="/images/icon_star.gif" border="0" align="absbottom">
			<b>PlusSale 추가구성 상품 등록</b>
		</td>
	</tr>
	<% if (oitem.FResultCount<1) then %>
	<tr height="25" bgcolor="FFFFFF">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
			<input type="button" class="button" value="검색" onClick="document.frm.submit();">
		</td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td colspan="3" align="center">[검색 결과가 없습니다.]</td>
	</tr>
	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
			<input type="button" class="button" value="검색" onClick="document.frm.submit();">
		</td>
		<td rowspan="4" width="100" align="right">
		    <img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100">
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td><%= oitem.FOneItem.FItemName %></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
		<td><%= oitem.FOneItem.FMakerid %></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">소비자가/매입가</td>
		<td>
		    <% if (oitem.FOneItem.FsaleYn="Y") then %>
    			<%= FormatNumber(oitem.FOneItem.FOrgPrice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.Forgsuplycash,oitem.FOneItem.FOrgPrice,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<br>

    			<font color=#F08050>(할)<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %></font>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<% if (oitem.FOneItem.IsCouponItem) then %>
    			<br><font color=#10F050>(쿠) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %></font>
    			<% end if %>
			<% else %>
    			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
    			&nbsp;
    			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
    			&nbsp;&nbsp;
    			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

    			<% if (oitem.FOneItem.IsCouponItem) then %>
    			<br><font color=#10F050>(쿠) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %> <!-- / <%= FormatNumber(oitem.FOneItem.Fcouponbuyprice) %> --> &nbsp;<%= oitem.FOneItem.GetCouponDiscountStr %> 할인 </font>
    			<% end if %>
			<% end if %>
		</td>
	</tr>
	<% end if %>
	</form>
</table>

<p>

<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmPlusSale" method="post" action="PlusDIYItem_Process.asp">
    <input type="hidden" name="osellcash" value="<%= oitem.FOneItem.FSellcash %>">
    <input type="hidden" name="obuycash" value="<%= oitem.FOneItem.FBuycash %>">
    <input type="hidden" name="omwdiv" value="<%= oitem.FOneItem.FMwDiv %>">
    <input type="hidden" name="itemid" value="<%= itemid %>">
    <% if (IsNewReg) then %>
    <input type="hidden" name="mode" value="regPlusSale">
    <% else %>
    <input type="hidden" name="mode" value="editPlusSale">
    <% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">플러스할인율</td>
		<td>
		    <% if IsNewReg then %>
		    할인율 :<input type="text" name="PlusSalePro" value="" class="text" size="5" maxlength="3" onKeyUp="CalcuMargin(frmPlusSale)">%
		    <% else %>
		    할인율 :<input type="text" name="PlusSalePro" value="<%= oPlusSaleItem.FOneItem.FPlusSalePro %>" class="text" size="5" maxlength="3" onKeyUp="CalcuMargin(frmPlusSale)">%
		    <% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">할인시공급율</td>
	    <td>
	        <% if IsNewReg then %>
    			<!-- 공급가설정 : -->
    			<select class="select" name="PlusSaleMaginFlag" onChange="CalcuMargin(frmPlusSale)">
    			    <option value="1" >동일마진</option>
                	<option value="2" >업체부담</option>
                	<!-- <option value="3" >반반부담</option> -->
                	<option value="4" >텐바이텐부담</option>
                	<option value="5" >직접설정</option>
    			</select>

    			<input type="text" name="PlusSaleMargin" class="text_ro" size="4" maxlength="4" onKeyUp="CalcuMargin(frmPlusSale)">%
    			&nbsp;&nbsp;
    			<input type="text" name="tmpSellCash" class="text_ro" size="10" maxlength="10" ReadOnly > / <input type="text" name="tmpBuyCash" value="" class="text_ro" size="10" maxlength="10" ReadOnly >
			<% else %>
			    <table border="0" cellspacing="0" cellpadding="0" class="a" >
			    <tr>
			        <td width="100" >
			            <select class="select" name="PlusSaleMaginFlag"  onChange="CalcuMargin(frmPlusSale)">
            			    <option value="1" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="1","selected","") %> >동일마진</option>
                        	<option value="2" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="2","selected","") %> >업체부담</option>
                        	<!-- <option value="3" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="3","selected","") %> >반반부담</option> -->
                        	<option value="4" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="4","selected","") %> >텐바이텐부담</option>
                        	<option value="5" <%= ChkIIF(oPlusSaleItem.FOneItem.FPlusSaleMaginFlag="5","selected","") %> >직접설정</option>
            			</select>
			        </td>
			        <td width="70">
			            <input type="text" name="PlusSaleMargin" class="text_ro" value="<%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>" size="4" maxlength="4" onKeyUp="CalcuMargin(frmPlusSale)">%
			        </td>
			        <td>
			            <input type="text" name="tmpSellCash" value="<%= oPlusSaleItem.FOneItem.getPlusSalePrice %>" class="text_ro" size="10" maxlength="10"> / <input type="text" name="tmpBuyCash" value="<%= oPlusSaleItem.FOneItem.getPlusSaleBuycash %>" class="text_ro" size="10" maxlength="10">
			        </td>
			    </tr>
			    <tr>
			        <td><%= oPlusSaleItem.FOneItem.getMaginFlagName %></td>
			        <td><%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>%</td>
			        <td>
			            <%= oPlusSaleItem.FOneItem.getPlusSalePrice %>
			            /
			            <%= oPlusSaleItem.FOneItem.getPlusSaleBuycash %>
			        </td>
			    </tr>
			    </table>
			<% end if %>
	    </td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">기간진행여부</td>
		<td>
		    <% if IsNewReg then %>
			<input type="radio" name="termsGubun" value="A" onClick="setComp(this);">상시진행
			<input type="radio" name="termsGubun" value="S" checked onClick="setComp(this);">기간진행
			<% else %>
			<input type="radio" name="termsGubun" value="A" <%= ChkIIF(oPlusSaleItem.FOneItem.IsAlwaysTerms,"checked","") %> onClick="setComp(this);">상시진행
			<input type="radio" name="termsGubun" value="S" <%= ChkIIF(oPlusSaleItem.FOneItem.IsAlwaysTerms,"","checked") %> onClick="setComp(this);">기간진행
			<% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">시작일</td>
		<td>
		    <% if IsNewReg then %>
		    <input type="text" class="text" name="PlusSaleStartDate" value="" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleStartDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (기간진행일 경우, 설정가능)
		    <% else %>
		    <input type="text" class="text" name="PlusSaleStartDate" value="<%= ChkIIF(Not IsNewReg,Left(oPlusSaleItem.FOneItem.FPlusSaleStartDate,10),"") %>" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleStartDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (기간진행일 경우, 설정가능)
		    <% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">종료일</td>
		<td>
		    <% if IsNewReg then %>
		    <input type="text" class="text" name="PlusSaleEndDate" value="" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (기간진행일 경우, 설정가능)
		    <% else %>
		    <input type="text" class="text" name="PlusSaleEndDate" value="<%= ChkIIF(Not IsNewReg,Left(oPlusSaleItem.FOneItem.FPlusSaleEndDate,10),"") %>" size="11" maxlength="10"> <a href="javascript:calendarOpen(frmPlusSale.PlusSaleEndDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=20></a> (기간진행일 경우, 설정가능)
		    <% end if %>
		</td>
	</tr>



	<!-- DB에 있는 상품만 보이는 메뉴 -->
	<% if (IsNewReg) then %>

	<% else %>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">현재상태</td>
		<td>
			 <!-- 진행예정 / 진행중 / 기간종료 (기간진행여부 및 기간으로 판단) -->
			 <%= oPlusSaleItem.FOneItem.getCurrstateName %>
		</td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td bgcolor="<%= adminColor("tabletop") %>">등록된 메인상품</td>
		<td>
			<%= oPlusSaleItem.FOneItem.FLinkedItemCount %> 개
			<input type="button" class="button" value="링크상품리스트" onclick="showLinkedItemList('<%= itemid %>');">
		</td>
	</tr>
	<% end if %>
	<!-- DB에 있는 상품만 보이는 메뉴 -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2" align="center">
		    <% if (IsNewReg) then %>
			<input type="button" class="button" value="신규등록" <%= ChkIIF(IsLinkedItem,"disabled","") %> onClick="RegPLusSale(frmPlusSale)";>
			<% else %>
			<input type="button" class="button" value=" 수  정 " onClick="RegPLusSale(frmPlusSale)";>
			&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value=" 삭  제 " onClick="DelPLusSale(frmPlusSale)";>
			<% end if %>
		</td>
	</tr>
	</form>
</table>
<% end if %>
<!--
<p>

상품코드 검색시 플러스세일상품 DB에 없을경우, 신규등록<br>
DB에 있을경우, 수정버튼 표시<br>
상품검색시 그 상품코드가 메인상품으로 사용되고 있을경우, 아래쪽 내용 대신 등록불가함을 표시<br>
<br>
  동일마진: 판매가 대비 동일 마진율 적용<br>
  업체부담: 원판매가의 마진금액만큼 할인판매가에서 차감 <br>
  반반부담: 할인금액의 1/2금액을 원공급가에서 차감<br>
  텐바이텐부담: 원공급가를 할인판매공급가로 고정 <br>
-->
<script language='javascript'>
function getOnLoad(){
    <% if (oitem.FResultCount>0) then %>
    <% if (oitem.FOneItem.FsaleYn="Y") then %>
    alert('이미 할인중인 상품입니다.');
    <% end if %>

    <% if (IsLinkedItem) then %>
    alert('이미 메인 링크에 등록된 상품입니다. - 플러스 세일 상품으로 등록 불가.');
    <% end if %>
    <% end if %>
}
window.onload = getOnLoad;
</script>
<%
set oitem = Nothing
set oitemoption = Nothing
set oPlusSaleItem = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
