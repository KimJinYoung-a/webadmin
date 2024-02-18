<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual ="/lib/classes/partners/partnerusercls.asp" -->
<%

dim itemid, oitem
dim makerid

itemid = requestCheckvar(request("itemid"),10)
makerid = requestCheckvar(request("makerid"),32)
menupos = RequestCheckvar(request("menupos"),10)
if (itemid = "") then
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbACADEMYget.close()	:	response.End
end if


'==============================================================================
set oitem = new CItem

oitem.FRectItemID = itemid
oitem.GetOneItem

'==============================================================================
''업체 기본계약 구분
dim defaultmargin, defaultmaeipdiv, defaultFreeBeasongLimit, defaultDeliverPay, defaultDeliveryType
Dim npartner, i
set npartner = new CPartnerUser
npartner.FRectDesignerID = oitem.FOneItem.Fmakerid
npartner.GetAcademyPartnerList
	defaultmargin			= npartner.FPartnerList(0).Fdiy_margin
    defaultmaeipdiv         = npartner.FPartnerList(0).Fmaeipdiv
    defaultFreeBeasongLimit = npartner.FPartnerList(0).FdefaultFreeBeasongLimit
    defaultDeliverPay       = npartner.FPartnerList(0).FdefaultDeliverPay
    defaultDeliveryType     = npartner.FPartnerList(0).FdefaultDeliveryType
set npartner = Nothing

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
   rsACADEMYget.Open query1,dbACADEMYget,1

   if  not rsACADEMYget.EOF  then
       rsACADEMYget.Movefirst

       do until rsACADEMYget.EOF
           if Lcase(selectedId) = Lcase(rsACADEMYget("userid")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsACADEMYget("userid")& "," & rsACADEMYget("defaultmargine") & "' "&tmp_str&">" & rsACADEMYget("userid") & "  [" & replace(db2html(rsACADEMYget("socname_kor")),"'","") & "]" & "</option>")
           tmp_str = ""
           rsACADEMYget.MoveNext
       loop
   end if
   rsACADEMYget.close
   response.write("</select>")
End Sub


%>
<script language="javascript" SRC="/js/confirm.js"></script>
<script language="javascript">

function UseTemplate() {
	window.open("/academy/comm/pop_basic_item_info_list.asp", "option_win", "width=600, height=420, scrollbars=yes, resizable=yes");
}

// ============================================================================
// 업체마진자동입력
function TnDesignerNMargineAppl(str){
	var varArray;
	varArray = str.split(',');

	document.itemreg.designerid.value = varArray[0];
	document.itemreg.margin.value = varArray[1];

}

function CalcuAuto(frm){
	var isvatYn, imileage;
	var isellcash, ibuycash, imargin;
	var isailprice, isailsuplycash, isailpricevat, isailsuplycashvat, isailmargin;

    isvatYn = frm.vatYn[0].checked;

	if (frm.saleYn[0].checked == true) {
	    // 정상가격
	    isellcash = frm.sellcash.value;
	    imargin = frm.margin.value;

    	if (imargin.length<1){
    		alert('마진을 입력하세요.');
    		frm.margin.focus();
    		return;
    	}

    	if (isellcash.length<1){
    		alert('판매가를 입력하세요.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (!IsDouble(imargin)){
    		alert('마진은 숫자로 입력하세요.');
    		frm.margin.focus();
    		return;
    	}

    	if (!IsDigit(isellcash)){
    		alert('판매가는 숫자로 입력하세요.');
    		frm.sellcash.focus();
    		return;
    	}

    	if (isvatYn==true){
    		ibuycash = isellcash - parseInt(isellcash*imargin/100);
    		imileage = parseInt(isellcash*0.01) ;
    	}else{
    		ibuycash = isellcash - parseInt(isellcash*imargin/100);
    		imileage = parseInt(isellcash*0.01) ;
    	}

    	frm.buycash.value = ibuycash;
    	frm.mileage.value = imileage;
	} else {
	    // 세일가격
	    isailprice = frm.sailprice.value;
	    isailmargin = frm.sailmargin.value;

    	if (isailmargin.length<1){
    		alert('세일마진을 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (isailprice.length<1){
    		alert('세일판매가를 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (!IsDouble(isailmargin)){
    		alert('세일마진은 숫자로 입력하세요.');
    		frm.sailmargin.focus();
    		return;
    	}

    	if (!IsDigit(isailprice)){
    		alert('세일판매가는 숫자로 입력하세요.');
    		frm.sailprice.focus();
    		return;
    	}

    	if (isvatYn==true){
    		isailpricevat = parseInt(parseInt(1/11 * parseInt(isailprice)));
    		isailsuplycash = isailprice - parseInt(isailprice*isailmargin/100);
    		isailsuplycashvat = parseInt(parseInt(1/11 * parseInt(isailsuplycash)));
    		imileage = parseInt(isailprice*0.01) ;
    	}else{
    		isailpricevat = 0;
    		isailsuplycash = isailprice - parseInt(isailprice*isailmargin/100);
    		isailsuplycashvat = 0;
    		imileage = parseInt(isailprice*0.01) ;
    	}

    	frm.sailpricevat.value = isailpricevat;
    	frm.sailsuplycash.value = isailsuplycash;
    	frm.sailsuplycashvat.value = isailsuplycashvat;
    	frm.mileage.value = imileage;
    }

	//할인율 계산
	if (frm.saleYn[0].checked == true) {
		document.getElementById("lyrPct").innerHTML = "";
	} else {
		isellcash = frm.sellcash.value;
		isailprice = frm.sailprice.value;
		var isalePercent = parseInt(Math.round((isellcash-isailprice)/isellcash*1000))/10;
		document.getElementById("lyrPct").innerHTML = "할인율: <font color='#EE0000'><strong>" + isalePercent + "%</strong></font>";
	}
}

// ============================================================================
// 저장하기
function SubmitSave() {
	if (itemreg.designerid.value == ""){
		alert("업체를 선택하세요.");
		itemreg.designer.focus();
		return;
	}

    if (validate(itemreg)==false) {
        return;
    }

    if (itemreg.saleYn[0].checked == true) {
        // 정상가격
        if (parseInt((itemreg.sellcash.value*1) * (itemreg.margin.value*1) / 100) != ((itemreg.sellcash.value*1) - (itemreg.buycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[소비자가*마진 = 공급가]");
    		itemreg.sellcash.focus();

    		if (!confirm('마진율로 계산 할 수 없을때 공급가만 입력하면 마진율은 공급가에 맞춰 계산됩니다. \n계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sellcash.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            itemreg.mileage.focus();
            return;
        }

        if (itemreg.sellcash.value*1 < 300 || itemreg.sellcash.value*1 >= 20000000){
			alert("판매 가격은 300원 이상 20,000,000만원 미만으로 등록 가능합니다.");
			itemreg.sellcash.focus();
			return;
		}

    } else {
        // 할인가격
        if (parseInt((itemreg.sailprice.value*1) * (itemreg.sailmargin.value*1) / 100) != ((itemreg.sailprice.value*1) - (itemreg.sailsuplycash.value*1))) {
    		alert("공급가가 잘못입력되었습니다.[할인소비자가*할인마진 = 할인공급가]");
    		itemreg.sailprice.focus();

    		if (!confirm('계속 진행 하시겠습니까?')){
				return;
			}
        }

        if (itemreg.mileage.value*1 > itemreg.sailprice.value*1){
            alert("마일리지는 판매가보다 클 수 없습니다.");
            itemreg.mileage.focus();
            return;
        }

        if (itemreg.sailprice.value*1 < 300 || itemreg.sailprice.value*1 >= 20000000){
			alert("판매 가격은 300원 이상 20,000,000만원 미만으로 등록 가능합니다.");
			itemreg.sailprice.focus();
			return;
		}
    }


    //세일가격이 정상가격 보다 클 수 없음.
    if (itemreg.sailprice.value*1>itemreg.sellcash.value*1){
        alert('세일가격이 정상가보다 클 수 없습니다.');
        return;
    }

    if (itemreg.sailsuplycash.value*1>itemreg.buycash.value*1){
        alert('세일매입가가 정상 매입가보다 클 수 없습니다.');
        return;
    }

    //배송구분 체크 =======================================
    //업체 조건배송
    if (!( ((itemreg.defaultFreeBeasongLimit.value*1>0) && (itemreg.defaultDeliverPay.value*1>0))||(itemreg.defaultDeliveryType.value=="9") )){
        if (itemreg.deliverytype[0].checked){
            alert('배송 구분을 확인해주세요. 개별배송 업체가 아닙니다.');
            return;
        }
    }

    //업체착불배송 : 조건배송도 착불설정가능
    if (!((itemreg.defaultDeliveryType.value=="7")||(itemreg.defaultDeliveryType.value=="9"))&&(itemreg.deliverytype[2].checked)){
        alert('배송 구분을 확인해주세요. [업체 착불배송,업체 조건배송] 업체가 아닙니다.');
        return;
    }

    //==================================================================================

    if(confirm("상품을 올리시겠습니까?") == true){
        itemreg.deliverytype[0].disabled=false;
		itemreg.deliverytype[1].disabled=false;
		itemreg.deliverytype[2].disabled=false;
        itemreg.submit();
    }

}


function TnChecksaleYn(frm){
	CheckSailEnDisabled(frm);
    CalcuAuto(frm);
}

function CheckSailEnDisabled(frm){
	if (frm.saleYn[0].checked == true) {
	    // 정상가격
        frm.sellcash.readonly = false;
        frm.margin.readonly = false;

        frm.sellcash.style.background = '#FFFFFF';
        frm.buycash.style.background = '#FFFFFF';
        frm.margin.style.background = '#FFFFFF';

        frm.sailprice.readonly = true;
        frm.sailmargin.readonly = true;

        frm.sailprice.style.background = '#E6E6E6';
        frm.sailsuplycash.style.background = '#E6E6E6';
        frm.sailmargin.style.background = '#E6E6E6';
	} else {
	    // 세일가격
        frm.sellcash.readonly = true;
        frm.margin.readonly = true;

        frm.sellcash.style.background = '#E6E6E6';
        frm.buycash.style.background = '#E6E6E6';
        frm.margin.style.background = '#E6E6E6';

        frm.sailprice.readonly = false;
        frm.sailmargin.readonly = false;

        frm.sailprice.style.background = '#FFFFFF';
        frm.sailsuplycash.style.background = '#FFFFFF';
        frm.sailmargin.style.background = '#FFFFFF';
    }
}

function ClearVal(comp){
    comp.value = "";
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
		<font color="red"><strong>상품 가격/판매 정보 수정</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br><b>등록된 상품의 가격 및 판매 정보를 수정합니다.</b>
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


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
  <form name="itemreg" method="post" action="itemmodify_Process.asp" onsubmit="return false;">
  <input type="hidden" name="mode" value="ItemPriceInfo">
  <input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
  <input type="hidden" name="designerid" value="<%= oitem.FOneItem.Fmakerid %>">

  <!-- 업체 기본 계약 구분 -->
  <input type="hidden" name="defaultmargin" value="<%= defaultmargin %>">
  <input type="hidden" name="defaultmaeipdiv" value="<%= defaultmaeipdiv %>">
  <input type="hidden" name="defaultFreeBeasongLimit" value="<%= defaultFreeBeasongLimit %>">
  <input type="hidden" name="defaultDeliverPay" value="<%= defaultDeliverPay %>">
  <input type="hidden" name="defaultDeliveryType" value="<%= defaultDeliveryType %>">

  <input type="hidden" name="orgprice" value="<%= oitem.FOneItem.Forgprice %>">
  <input type="hidden" name="orgsuplycash" value="<%= oitem.FOneItem.Forgsuplycash %>">
  <input type="hidden" name="availPayType" value="<%= oitem.FOneItem.FavailPayType %>">
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품코드 :</td>
  	<td bgcolor="#FFFFFF" colspan="3">
  	  <%= oitem.FOneItem.Fitemid %>
  	  &nbsp;&nbsp;&nbsp;&nbsp;
  	  <input type="button" value="미리보기" onclick="window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>');">
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">업체ID :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%=oitem.FOneItem.FMakerid %>&nbsp;&nbsp;(마진 : <%= defaultmargin %>%)</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">상품명 :</td>
  	<td bgcolor="#FFFFFF" colspan="3"><%= oitem.FOneItem.Fitemname %></td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>가격정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">

  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">가격설정 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
      <table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
        <tr align="center">
          <td height="25" width="90" bgcolor="#DDDDFF">선택</td>
          <td width="100" bgcolor="#DDDDFF">소비자가</td>
          <td width="100" bgcolor="#DDDDFF">공급가</td>
          <td width="100" bgcolor="#DDDDFF">마진</td>
          <td bgcolor="#DDDDFF">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" bgcolor="#FFFFFF"><input type="radio" name="saleYn" onClick="TnChecksaleYn(itemreg)" value="N" <% if oitem.FOneItem.FsaleYn = "N" then response.write "checked" %>> 정상가격</td>
          <td bgcolor="#FFFFFF" align="center">
            <% if oitem.FOneItem.FsaleYn = "N" then %>
            <input type="text" name="sellcash" maxlength="16" size="8" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Fsellcash %>" onkeyup="CalcuAuto(itemreg);">원
            <% else %>
            <input type="text" name="sellcash" maxlength="16" size="8" id="[on,on,off,off][소비자가]" value="<%= oitem.FOneItem.Forgprice %>" onkeyup="CalcuAuto(itemreg);">원
            <% end if %>
          </td>
          <td bgcolor="#FFFFFF" align="center">
            <% if oitem.FOneItem.FsaleYn = "N" then %>
            <input type="text" name="buycash" maxlength="16" size="8" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fbuycash %>">원
            <% else %>
            <input type="text" name="buycash" maxlength="16" size="8" id="[on,on,off,off][공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Forgsuplycash %>">원
            <% end if %>
          </td>
          <% if oitem.FOneItem.FsaleYn = "N" then %>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][마진]" value="<%= margine %>">%
          </td>
          <% else %>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="margin" maxlength="32" size="5" id="[on,off,off,off][마진]" value="<%= orgmargine %>">%
          </td>
          <% end if %>
          <td bgcolor="#FFFFFF" align="left">
            <input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);">
          </td>
        </tr>
        <tr>
          <td height="25" bgcolor="#FFFFFF"><input type="radio" name="saleYn" onClick="TnChecksaleYn(itemreg)" value="Y" <% if oitem.FOneItem.FsaleYn = "Y" then response.write "checked" %>> 할인가격</td>
          <input type="hidden" name="sailpricevat">
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailprice" maxlength="16" size="8" id="[on,on,off,off][할인소비자가]" value="<%= oitem.FOneItem.Fsailprice %>"  onkeyup="CalcuAuto(itemreg);">원
          </td>
          <input type="hidden" name="sailsuplycashvat">
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailsuplycash" maxlength="16" size="8" id="[on,on,off,off][할인공급가]"  style="background-color:#E6E6E6;" value="<%= oitem.FOneItem.Fsailsuplycash %>">원
          </td>
          <td bgcolor="#FFFFFF" align="center">
            <input type="text" name="sailmargin" maxlength="32" size="5" id="[on,off,off,off][할인마진]" value="<%= sailmargine %>">%
          </td>
          <td bgcolor="#FFFFFF" align="left">
            <input type="button" value="공급가 자동계산" onclick="CalcuAuto(itemreg);">
			<%
				dim itemSalePer : itemSalePer=0
				if oitem.FOneItem.FsaleYn="Y" then
					itemSalePer = oitem.FOneItem.Forgprice - oitem.FOneItem.Fsailprice
					itemSalePer = itemSalePer/oitem.FOneItem.Forgprice*100
				end if
			%>
			<span id="lyrPct" style="white-space:nowrap;"><% if itemSalePer>0 then %>할인율: <font color="#EE0000"><strong><%=formatNumber(itemSalePer,1)%>%</strong></font><% end if %></span>
          </td>
        </tr>
      </table>
      <br>
      - 공급가는 <b>부가세 포함가</b>입니다.<br>
      - 소비자가(할인가)와 마진(할인마진)을 입력하고 [공급가자동계산] 버튼을 누르면 공급가와 마일리지가 자동계산됩니다.
  	</td>
  </tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">마일리지 :</td>
  	<td width="35%" bgcolor="#FFFFFF"><input type="text" name="mileage" maxlength="32" size="10" id="[on,on,off,off][마일리지]" value="<%= oitem.FOneItem.Fmileage %>">point</td>
  	<td width="15%" bgcolor="#DDDDFF">과세, 면세 여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
      <input type="radio" name="vatYn" value="Y" <% if oitem.FOneItem.FvatYn = "Y" then response.write "checked" %>>과세
      <input type="radio" name="vatYn" value="N" <% if oitem.FOneItem.FvatYn = "N" then response.write "checked" %>>면세
  	</td>
  </tr>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
          <br>판매정보
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">매입특정구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<%= oitem.FOneItem.Fmwdiv %> <font color="red">**변경불가</font>
		<input type="hidden" name="mwdiv" value="<%= oitem.FOneItem.Fmwdiv %>">
  	</td>
</tr>
<tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">배송정책구분 :</td>
  	<td width="35%" bgcolor="#FFFFFF" colspan="3">
		<input type="radio" name="deliverytype" value="1" <% if oitem.FOneItem.Fdeliverytype = "1" then response.write "checked" %>>텐바이텐배송&nbsp;
		<input type="radio" name="deliverytype" value="2" <% if oitem.FOneItem.Fdeliverytype = "2" then response.write "checked" %>>업체(무료)배송&nbsp;
		<input type="radio" name="deliverytype" value="4" <% if oitem.FOneItem.Fdeliverytype = "4" then response.write "checked" %>>텐바이텐무료배송
      	<input type="radio" name="deliverytype" value="9" <% if oitem.FOneItem.Fdeliverytype = "9" then response.write "checked" %>>업체조건배송(개별 배송비부과)
  	  	<input type="radio" name="deliverytype" value="7" <% if oitem.FOneItem.Fdeliverytype = "7" then response.write "checked" %>>업체착불배송
  	</td>
</tr>
  <tr align="left">
  	<td height="30" width="15%" bgcolor="#DDDDFF">판매여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	  <input type="radio" name="sellyn" value="Y" <% if oitem.FOneItem.Fsellyn = "Y" then response.write "checked" %>>판매함&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="S" <% if oitem.FOneItem.Fsellyn = "S" then response.write "checked" %>>일시품절&nbsp;&nbsp;
  	  <input type="radio" name="sellyn" value="N" <% if oitem.FOneItem.Fsellyn = "N" then response.write "checked" %>>판매안함
  	</td>
  	<td height="30" width="15%" bgcolor="#DDDDFF">사용여부 :</td>
  	<td width="35%" bgcolor="#FFFFFF">
  	    <input type="radio" name="isusing" value="Y" <% if oitem.FOneItem.Fisusing = "Y" then response.write "checked" %>>사용함&nbsp;&nbsp;
  	    <input type="radio" name="isusing" value="N" <% if oitem.FOneItem.Fisusing = "N" then response.write "checked" %>>사용안함
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
<!-- 표 하단바 끝-->

<p>
<script language='javascript'>
// 매입특정구분 및 배송구분세팅
for (var i = 0; i < itemreg.elements.length; i++) {
    if (itemreg.elements(i).name == "deliverytype") {
        if (itemreg.elements(i).value == "<%= oitem.FOneItem.Fdeliverytype %>") {
            itemreg.elements(i).checked = true;
        }
    }
}

// 세일
CheckSailEnDisabled(itemreg);
</script>
<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
