<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%
dim itemid
itemid = getNumeric(requestCheckVar(request("itemid"),10))

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if itemid<>"" then
    if (NOT C_ADMIN_USER) then
    oitem.FRectMakerid = session("ssBctID")
    end if
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

dim i

%>
<script>

//한정 or 비한정 Radio버튼 클릭시
function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && (e.name.substring(0,"optremainno".length) == "optremainno")) {
				e.disabled = (comp.value=="N");
		  }
  	}

    if (comp.value=="N"){
        resetLimit2Zero();
    }
}

//한정수량 0으로 Setting
function resetLimit2Zero(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    e.value = 0;
		  	}
		}
  	}
}

function SaveItem(frm){
<% if oitem.FResultCount>0 then %>
    <% if Not oitem.FOneItem.IsUpchebeasong then %>
//핑거스 단종 없음?! 2016-07-12 주석처리
//    //판매 N 인경우 단종품절 또는 MD품절로 설정 해야함.
//    if ((frm.sellyn[2].checked)&&!((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked))){
//        alert('판매 중지 상품인경우 재고부족,단종품절 또는 MD품절로 설정하셔야 합니다.');
//        frm.danjongyn[2].focus();
//        return;
//    }
//
//    //재고부족,단종설정은 한정판매인경우만 가능함
//	if ((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked)){
//		if (!frm.limityn[0].checked){
//			alert('한정 판매인 경우만 재고부족,단종품절, MD품절로 설정 할 수 있습니다.');
//			frm.limityn[0].focus();
//			return;
//		}
//	}
	<% end if %>
<% end if %>

	//사용안함이나 전시하는경우
	if ((frm.isusing[1].checked)&&(frm.sellyn[0].checked)){
        alert('사용 중지 상품은 판매로 설정 불가합니다.');
        frm.sellyn[2].focus();
        return;
    }


	frm.itemoptionarr.value = "";
	//옵션 한정 남은 수량
	frm.optremainnoarr.value = "";
	//옵션 사용 여부
	frm.optisusingarr.value = "";

    var option_isusing_count = 0;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    //숫자만 가능
		  	    if (!IsDigit(e.value)){
		  	        alert('한정 수량은 숫자만 가능합니다.');
		  	        e.select();
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + e.id + "," ;
				frm.optremainnoarr.value = frm.optremainnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                }
		  	}

            //옵션 사용여부
			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		}
  	}

    if (option_isusing_count < 1) {
        alert("모든 옵션을 사용안함으로 할수 없습니다. 상품정보를 사용안함으로 변경하거나, 전시안함 변경하세요.");
        return;
    }


	var ret = confirm('저장 하시겠습니까?');

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/academy/comm/pop_diyitemoptionedit.asp?itemid=' + iid,'popitemoptionedit','width=700 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function ReloadWindow() {
    document.location.reload();
}

window.resizeTo(560,550);
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	상품코드 : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="7" size="7">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->



<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="do_diy_simpleiteminfoedit.asp">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoptionarr value="">
<input type=hidden name=optisusingarr value="">
<input type=hidden name=optremainnoarr value="">

	<tr>
	<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 border="0" class=a bgcolor=#BABABA>
			<tr height="25">

		<td width="120" bgcolor="#DDDDFF">상품명</td>
		<td colspan="2" bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">브랜드ID/브랜드명</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= oitem.FOneItem.Fmakerid %>/<%= oitem.FOneItem.FBrandName %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">소비자가/매입가</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			&nbsp;&nbsp;
			<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %> %
			<% end if %>
		</td>
	</tr>

	<% if (oitem.FOneItem.FsaleYn="Y") then %>
	<tr height="25">
		<td bgcolor="#DDDDFF">할인가/매입가</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<font color="red">
				<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
				&nbsp;&nbsp;
				<% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %>%
			    <% end if %>
			    할인
			</font>
			&nbsp;&nbsp;
			<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
				<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
		</td>
	</tr>
	<% end if %>

	<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>
	<tr height="25">
		<td bgcolor="#DDDDFF">쿠폰가/매입가</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<font color="green">
				<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>
				&nbsp;&nbsp;
				<%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			</font>
		</td>
	</tr>
	<% end if %>
	<tr height="25">
		<td bgcolor="#DDDDFF">사용옵션</td>
		<td bgcolor="#FFFFFF">
		(<%= oitem.FOneItem.FOptionCnt %> 개)
		&nbsp;
		<input type=button class="button" value="옵션추가/수정" onclick="popoptionEdit('<%= itemid %>');">
		</td>
		<td rowspan="3" align="right" bgcolor="#FFFFFF">
			<img src="<%= oitem.FOneItem.FListImage %>" width="100" align="right">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">배송구분</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsUpcheBeasong then %>
		<b>업체</b>배송
		<% else %>
		텐바이텐배송
		<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">상품 품절여부</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsSoldOut then %>
		<font color=red><b>품절</b></font>
		<% end if %>
		</td>
	</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">

			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">상품 판매여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FSellYn="Y" then %>
					<input type="radio" name="sellyn" value="Y" checked >판매함
					<input type="radio" name="sellyn" value="S" >일시품절
					<input type="radio" name="sellyn" value="N" >판매안함
					<% elseif oitem.FOneItem.FSellYn="S" then %>
					<input type="radio" name="sellyn" value="Y" >판매함
					<input type="radio" name="sellyn" value="S" checked ><font color="red">일시품절</font>
					<input type="radio" name="sellyn" value="N" >판매안함
					<% else %>
					<input type="radio" name="sellyn" value="Y" >판매함
					<input type="radio" name="sellyn" value="S" >일시품절
					<input type="radio" name="sellyn" value="N" checked ><font color="red">판매안함</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">상품 사용여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FIsUsing="Y" then %>
					<input type="radio" name="isusing" value="Y" checked >사용함
					<input type="radio" name="isusing" value="N" >사용안함
					<% else %>
					<input type="radio" name="isusing" value="Y" >사용함
					<input type="radio" name="isusing" value="N" checked ><font color="red">사용안함</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">한정판매여부</td>
				<td bgcolor="#FFFFFF">
				<% if oitem.FOneItem.FLimitYn="Y" then %>
				<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">한정판매</font>
				<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">비한정판매
				(<%= oitem.FOneItem.FLimitNo %>-<%= oitem.FOneItem.FLimitSold %>=<%= oitem.FOneItem.FLimitNo-oitem.FOneItem.FLimitSold %>)
				<% else %>
				<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">한정판매
				<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">비한정판매
				<% end if %>
				</td>
			</tr>
			</table>

		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25" align="center" bgcolor="#FFDDDD" >
				<td width="50">옵션코드</td>
				<td>옵션명</td>
				<td width="100">옵션사용여부</td>
				<td width="40">현재<br>한정</td>
				<td width="80">한정판매수량</td>
			</tr>
			<% if oitemoption.FResultCount>0 then %>
				<% for i=0 to oitemoption.FResultCount - 1 %>
					<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
					<tr align="center" bgcolor="#EEEEEE">
					<% else %>
					<tr align="center" bgcolor="#FFFFFF">
					<% end if %>
						<td><%= oitemoption.FITemList(i).FItemOption %></td>
						<td><%= oitemoption.FITemList(i).FOptionName %></td>
						<td>
							<% if oitemoption.FITemList(i).FOptIsUsing="Y" then %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >N
							<% else %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">N</font>
							<% end if %>
						</td>
						<td><%= oitemoption.FITemList(i).GetOptLimitEa %></td>
						<td>
							<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
					</tr>
				<% next %>
			<% else %>
					<tr align="center" bgcolor="#FFFFFF">
						<td>0000</td>
						<td colspan="2">옵션없음</td>
						<td><%= oitem.FOneItem.GetLimitEa %></td>
						<td>
							<input type="text" id="0000" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
					</tr>
			<% end if %>
			</table>
		</td>
	</tr>
</form>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
          <input type="button" value="저장하기" onclick="SaveItem(frm2)">
          <input type="button" value=" 닫 기 " onclick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<% else %>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
    <td align="center">[검색 결과가 없습니다.]</td>
</tr>
</table>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
