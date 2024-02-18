<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/new_itemcls.asp"-->
<%
dim itemid
itemid = request("itemid")

dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

dim i

''브랜드 랙코드
dim sqlStr, prtidx
if (itemid<>"") then
    sqlStr = "select prtidx from [db_user].[10x10].tbl_user_c "
    sqlStr = sqlStr & " where userid='" & oitem.FOneItem.FMakerid & "'"
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        prtidx = rsget("prtidx")

        prtidx = format00(4,prtidx)
    end if
    rsget.close
end if

%>
<script language='javascript'>

//한정 or 비한정 Radio버튼 클릭시
function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && (e.name.substring(0,"optremainno".length) == "optremainno")) {
				e.disabled = (comp.value=="N");
		  }
  	}

    frm.recalcuLimit.disabled = (comp.value=="N");

    if (comp.value=="N"){
        resetLimit2Zero();
    }else{
        resetLimit();
    }
}

//한정수량 재설정
function resetLimit(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    //Enable 인 경우만
		  	    if (!e.disabled){
		  	        //현재 재고의 97%로 설정 (재고가 10개 이상인 경우만) 내림
		  	        if (e.dumistock>=10){
		  	            e.value = parseInt(e.dumistock*0.97);
		  	        }

		  	    }
		  	}
		}
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
	if ((frm.itemrackcode.value.length>0)&&(frm.itemrackcode.value.length!=4)){
		alert('상품 랙코드는 4자리로 고정되어있습니다.');
		frm.itemrackcode.focus();
		return;
	}

    //전시 판매 속성이 다를 수 없음: ->변경 전시O 판매X 가능,
    if (frm.dispyn[1].checked&&frm.sellyn[0].checked){
        alert('전시 안하는 상품을 판매 할 수 없습니다.');
        frm.dispyn[0].focus();
        return;
    }
<% if (session("ssBctId") <> "icommang") then %>
    if ((frm.dispyn[0].checked&&frm.sellyn[1].checked)||(frm.dispyn[1].checked&&frm.sellyn[0].checked)){
        alert('전시 판매 속성을 다르게 설정 할 수 없습니다. \n\n전시안함=판매안한 or 전시함=판매함');
        frm.dispyn[0].focus();
        return;
    }
<% end if %>

<% if oitem.FResultCount>0 then %>
    <% if Not oitem.FOneItem.IsUpchebeasong then %>
    //(전시) 판매 N 인경우 단종품절 또는 MD품절로 설정 해야함.
    if ((frm.sellyn[1].checked)&&!((frm.danjongyn[2].checked)||(frm.danjongyn[3].checked))){
        alert('판매 중지 상품인경우 단종품절 또는 MD품절로 설정하셔야 합니다.');
        frm.danjongyn[2].focus();
        return;
    }

    //일시품절,단종설정은 한정판매인경우만 가능함
	if ((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked)){
		if (!frm.limityn[0].checked){
			alert('한정 판매인 경우만 일시품절,단종품절, MD품절로 설정 할 수 있습니다.');
			frm.limityn[0].focus();
			return;
		}
	}
	<% end if %>
<% end if %>

	//사용안함이나 전시하는경우
	if ((frm.isusing[1].checked)&&(frm.dispyn[0].checked)){
        alert('사용 중지 상품은 판매로 설정 불가합니다.');
        frm.dispyn[1].focus();
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
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + iid,'popitemoptionedit','width=700 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function ReloadWindow() {
    document.location.reload();
}

window.resizeTo(560,700);
</script>

<!-- TOP -->
<table width="280" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
  <tr height="20">
  	<td>
    	<img src="/images/icon_star.gif" align="absbottom"><font color="red"><strong>상품검색</strong></font>
	</td>
	<td align="right">
    	<a href="/PDAadmin/index.asp">HOME</a>
	</td>
  </tr>
</table>
<!-- TOP -->

<!-- 표 상단검색 시작-->
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td valign="top" bgcolor="F4F4F4">
	        	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="13">
	        	<input type="button" class="button" value="검색">
	        </td>
	</tr>
	</form>
</table>
<!-- 표 상단검색 끝-->



<% if oitem.FResultCount>0 then %>
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="do_simpleiteminfoedit.asp">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoptionarr value="">
<input type=hidden name=optisusingarr value="">
<input type=hidden name=optremainnoarr value="">

	<tr bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.Fmakerid %>/<%= oitem.FOneItem.FBrandName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %>/<%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %>%
			<% end if %>
		</td>
	</tr>
	<% if (oitem.FOneItem.FSailYn="Y") then %>
	<tr  bgcolor="#FFFFFF">
		<td>
			<font color="red">
				<%= FormatNumber(oitem.FOneItem.FSellcash,0) %>/<%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
				<% if (oitem.FOneItem.Forgprice<>0) then %>
					할인
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %>%
			    <% end if %>

			</font>
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
			<% if oitem.FOneItem.FSellcash<>0 then %>
				<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %>%
			<% end if %>
		</td>
	</tr>
	<% end if %>

	<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>
	<tr bgcolor="#FFFFFF">
		<td>
			<font color="green">
				<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>
				&nbsp;&nbsp;
				<%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			</font>
		</td>
	</tr>
	<% end if %>

	<tr align="right" bgcolor="#FFFFFF">
		<td>
			상품랙코드<input type="text" class="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="4" maxlength="4" >
		</td>
	</tr>
</table>

<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr align="center" bgcolor="#FFFFFF">
		<td>전시</td>
		<td>판매</td>
		<td>사용</td>
		<td>단종</td>
		<td>한정</td>
	</td>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.FDispYn %></td>
		<td><%= oitem.FOneItem.FSellYn %></td>
		<td><%= oitem.FOneItem.FIsUsing %></td>
		<td><%= oitem.FOneItem.Fdanjongyn %></td>
		<td><%= oitem.FOneItem.FLimitYn %></td>
	</td>
</table>
<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
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
			&nbsp;&nbsp;
			브랜드랙코드 : <%= prtidx %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">소비자가/매입가</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			&nbsp;&nbsp;
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %> %
			<% end if %>
		</td>
	</tr>

	<% if (oitem.FOneItem.FSailYn="Y") then %>
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
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
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
		<td bgcolor="#DDDDFF">상품랙코드</td>
		<td bgcolor="#FFFFFF" width="270">
			<input type="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="4" maxlength="4" > (4자리 Fix)
		</td>
		<td rowspan="4" align="right" bgcolor="#FFFFFF">
			<img src="<%= oitem.FOneItem.FListImage %>" width="100" align="right">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">사용옵션</td>
		<td bgcolor="#FFFFFF">
		(<%= oitem.FOneItem.FOptionCnt %> 개)
		&nbsp;
		<input type=button class="button" value="옵션수정" onclick="popoptionEdit('<%= itemid %>');">
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
				<td width="120" bgcolor="#DDDDFF">상품 전시여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FDispYn="Y" then %>
					<input type="radio" name="dispyn" value="Y" checked >전시함
					<input type="radio" name="dispyn" value="N" >전시안함
					<% else %>
					<input type="radio" name="dispyn" value="Y" >전시함
					<input type="radio" name="dispyn" value="N" checked ><font color="red">전시안함</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">상품 판매여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FSellYn="Y" then %>
					<input type="radio" name="sellyn" value="Y" checked >판매함
					<input type="radio" name="sellyn" value="N" >판매안함
					<% else %>
					<input type="radio" name="sellyn" value="Y" >판매함
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
				<td bgcolor="#DDDDFF">상품 단종여부</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.Fdanjongyn="Y" then %>
    					<input type="radio" name="danjongyn" value="N" >생산중
    					<input type="radio" name="danjongyn" value="S" >일시품절(7일이상)
    					<input type="radio" name="danjongyn" value="Y" checked ><font color="red">단종품절</font>
    					<input type="radio" name="danjongyn" value="M" >MD품절
					<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
    					<input type="radio" name="danjongyn" value="N" >생산중
    					<input type="radio" name="danjongyn" value="S" checked ><font color="red">일시품절(7일이상)</font>
    					<input type="radio" name="danjongyn" value="Y" >단종품절
    					<input type="radio" name="danjongyn" value="M" >MD품절
					<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
    					<input type="radio" name="danjongyn" value="N" >생산중
    					<input type="radio" name="danjongyn" value="S" >일시품절(7일이상)
    					<input type="radio" name="danjongyn" value="Y" >단종품절
    					<input type="radio" name="danjongyn" value="M" checked ><font color="red">MD품절</font>
					<% else %>
    					<input type="radio" name="danjongyn" value="N" checked >생산중
    					<input type="radio" name="danjongyn" value="S" >일시품절(7일이상)
    					<input type="radio" name="danjongyn" value="Y" >단종품절
    					<input type="radio" name="danjongyn" value="M" >MD품절
					<% end if %>
					<font color="#AAAAAA">
					<br> (상품판매에는 영향없슴 - 추가 입고예정 없을시 단종설정)
				</font>
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
	    <td colspan="2" bgcolor="#FFFFFF">한정비교재고가 10미만일 경우는 재고파악 후 수기로 입력하시기 바랍니다.</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25" align="center" bgcolor="#FFDDDD" >
				<td width="50">옵션코드</td>
				<td>옵션명</td>
				<td width="100">옵션사용여부</td>
				<td width="40">현재<br>한정</td>
				<td width="80">한정판매수량<br><input name="recalcuLimit" type="button" class="button" value="한정재계산" onclick="resetLimit();" <%= chkIIF(oitem.FOneItem.FLimitYn="N","disabled","") %>></td>
				<td width="80"><a href="javascript:TnPopItemStock('<%= itemid %>','');">한정비교재고</a></td>
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
							<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >N
							<% else %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">N</font>
							<% end if %>
						</td>
						<td><%= oitemoption.FITemList(i).GetOptLimitEa %></td>
						<td>
							<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" dumistock="<%= oitemoption.FITemList(i).GetLimitStockNo %>" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitemoption.FITemList(i).GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','<%= oitemoption.FITemList(i).FItemOption %>');"><%= oitemoption.FITemList(i).GetLimitStockNo %></a></td>
					</tr>
				<% next %>
			<% else %>
					<tr align="center" bgcolor="#FFFFFF">
						<td>0000</td>
						<td colspan="2">옵션없음</td>
						<td><%= oitem.FOneItem.GetLimitEa %></td>
						<td>
							<input type="text" id="0000" dumistock="<%= oitem.FOneItem.GetLimitStockNo %>" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitem.FOneItem.GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','');"><%= oitem.FOneItem.GetLimitStockNo %></a></td>
					</tr>
			<% end if %>
			</table>
		</td>
	</tr>

	<input type=hidden name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">

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
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->