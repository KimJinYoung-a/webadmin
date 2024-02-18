<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%

// 97% -> 98%, 2014-07-25
const C_LIMIT_PERCENT = 1
'const C_LIMIT_PERCENT = 0.98
'//C_LIMIT_PERCENT(한정조정 비율 변경) 2019-01-03 정태훈
public Function GetDanJongStat(byval v)
	if v="Y" then
		GetDanJongStat="단종"
	elseif v="S" then
		GetDanJongStat="일시품절"
	elseif v="M" then
		GetDanJongStat="MD품절"
	elseif v="N" then
		GetDanJongStat="생산중"
	else
	    GetDanJongStat=v
	end if
End Function

dim idx
idx = request("idx")





dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
oipchul.FRectId = idx
oipchul.GetIpChulMaster

set oipchuldetail = new CIpChulStorage
oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
oipchuldetail.GetIpChulDetailCheck

dim i
dim maylimitno, mayreallimtno, currlimitea
dim addmayno
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function ReLimitSell(itemid,itemoption,maylimitno,mayreallimtno,currlimitno){
	if (confirm('입고시 예상재고: ' + maylimitno + '\n\n한정수량 [   ' + mayreallimtno + '    ] \n\n로 조정하시겠습니까?\n\n(현재:' + currlimitno + '  추가:'+(mayreallimtno-currlimitno) + ')')){
		var popwin = window.open('/admin/newstorage/ipgoitemlimitcheckNew_process.asp?mode=addmaylimit&itemid=' + itemid + '&itemoption=' + itemoption + '&mayreallimtno=' + mayreallimtno,'doipgoitemlimitcheck','width=100 height=100');
	}
}

function ReLimitSellIMSI(itemid,itemoption,oldno,addno){
	if (confirm('입고시 한정수량: \n\ 한정수량 [   ' + (oldno*1 + addno*1) + '    ] \n\n로 조정하시겠습니까?\n\n(현재:' + oldno + '  추가:'+(addno) + ')')){
		var popwin = window.open('/admin/newstorage/doipgoitemlimitcheck.asp?mode=addlimitno&itemid=' + itemid + '&itemoption=' + itemoption + '&addno=' + addno,'doipgoitemlimitcheck','width=100 height=100');
	}
}

// ============================================================================
// 옵션수정
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnEditRealStockNo(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('변경할 한정 비교재고가 없습니다.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.arrCheckinfo.value = upfrm.arrCheckinfo.value + "|" + frm.cksel.value;
				upfrm.arrStockNo.value = upfrm.arrStockNo.value + "|" + frm.mayreallimtno.value;
				//upfrm.arrLimitYN.value = upfrm.arrLimitYN.value + "|" + frm.limityn.value;
			}
		}
	}
	upfrm.submit();
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	총건수:  <%= oipchuldetail.FResultCount %>
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="button" class="button" value="선택 한정조정" onClick="fnEditRealStockNo();">&nbsp;&nbsp;<input type="button" class="button" value="새로고침" onClick="document.location.reload();">
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
	<tr bgcolor="#DDDDFF" align="center">
		<td width="50"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="120">상품코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="50">단종여부</td>
		<td width="50">판매가</td>
		<td width="40">주문<br>수량</td>
		<td width="40">입고<br>수량</td>
		<td width="50">한정<br>비교재고</td>
		<td width="30">품절</td>
		<td width="30">판매</td>
		<td width="30">한정</td>
		<td width="50">한정<br>수량</td>
		<td width="30">한정<br>조정</td>
		<td width="50">비고 </td>
	</tr>
	<% for i=0 to oipchuldetail.FResultCount-1 %>
	<%
		''if (oipchuldetail.FItemList(i).FIsNewItem = "Y") then
		''    maylimitno = oipchuldetail.FItemList(i).FMaystockno + oipchuldetail.FItemList(i).Fitemno ''(입고수량)
		''else
		    maylimitno = oipchuldetail.FItemList(i).FMaystockno
		''end if

		mayreallimtno = Fix(maylimitno*C_LIMIT_PERCENT)

		currlimitea = oipchuldetail.FItemList(i).getOptionLimitEa

	%>
	<form name="frmBuyPrc_<%= oipchuldetail.FItemList(i).Fiitemgubun %>_<%= oipchuldetail.FItemList(i).FItemId %>_<%= oipchuldetail.FItemList(i).FItemOption %>" method="post" >
	<input type="hidden" name="mayreallimtno" value="<%=mayreallimtno%>">
	<tr align=center bgcolor="<%= oipchuldetail.FItemList(i).GetMayCheckColor %>">
		<td height="25">
		<% if (oipchuldetail.FItemList(i).FLimitYn="Y") then %>
		    <% if mayreallimtno<>0 then %>
        		<% if ((mayreallimtno<>currlimitea) or (currlimitea>mayreallimtno)) then %>
				<input type="checkbox" name="cksel" id="cksel" value="<%= oipchuldetail.FItemList(i).Fiitemgubun %>_<%= oipchuldetail.FItemList(i).FItemId %>_<%= oipchuldetail.FItemList(i).FItemOption %>">
				<% else %>
				<input type="checkbox" name="cksel" id="cksel" disabled>
		        <% end if %>
			<% else %>
			<input type="checkbox" name="cksel" id="cksel" disabled>
		    <% end if %>
		<% else %>
			<input type="checkbox" name="cksel" id="cksel" disabled>
		<% end if %>
		</td>
		<td>
			<font color="<%= mwdivColor(oipchuldetail.FItemList(i).FOnlineMwdiv) %>"><%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= oipchuldetail.FItemList(i).FItemId %>-<%= oipchuldetail.FItemList(i).FItemOption %></font>
		</td>
		<td align="left">
		    <a href="javascript:PopItemSellEdit('<%= oipchuldetail.FItemList(i).FItemID %>');"><%= oipchuldetail.FItemList(i).Fiitemname %></a>
		</td>
		<td <%= chkIIF(oipchuldetail.FItemList(i).FItemOption<>"0000" and oipchuldetail.FItemList(i).FOptUsing="N","bgcolor='#FF3333'","") %>>
		    <%= oipchuldetail.FItemList(i).Fiitemoptionname %>
		</td>
		<td><%= fncolor(oipchuldetail.FItemList(i).FDanJongYn,"dj") %></td>
		<td align=right><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
		<td><%= oipchuldetail.FItemList(i).Fbaljuitemno %></td>
		<td><%= oipchuldetail.FItemList(i).Fitemno %></td>
        <% if (oipchuldetail.FItemList(i).FIsNewItem = "Y") then %>
        <td><%= oipchuldetail.FItemList(i).FMaystockno %></td>
        <% else %>
        <td><%= oipchuldetail.FItemList(i).FMaystockno %></td>
        <% end if %>

		<td><font color="<%= oipchuldetail.FItemList(i).GetIsSlodOutColor %>"><%= oipchuldetail.FItemList(i).GetIsSlodOutText %></font></td>
		<td><font color="<%= oipchuldetail.FItemList(i).GetSellYnColor %>"><%= oipchuldetail.FItemList(i).Fsellyn %></font></td>

		<% if oipchuldetail.FItemList(i).Flimityn="Y" then %>
		<td bgcolor="#FFDDDD"><font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font></td>
		<% else %>
		<td><font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font></td>
		<% end if %>

		<td>
		<% if oipchuldetail.FItemList(i).FLimitYn="Y" then %>
			<%= oipchuldetail.FItemList(i).getOptionLimitEa %>
		<% end if %>
		</td>
		<td>
		<%=mayreallimtno%>/<%=currlimitea%>
		<% if (oipchuldetail.FItemList(i).FLimitYn="Y") then %>
		    <% if mayreallimtno<>0 then %>
        		<% if ((mayreallimtno<>currlimitea) or (currlimitea>mayreallimtno)) then %>
        		<input type="button" class="button" value="->" onclick="ReLimitSell('<%= oipchuldetail.FItemList(i).FItemID %>','<%= oipchuldetail.FItemList(i).FItemOption %>','<%= maylimitno %>','<%= mayreallimtno %>','<%= currlimitea %>')">
        		<% end if %>
		    <% end if %>
		<% end if %>
		</td>
		<td><% if oipchuldetail.FItemList(i).FDtComment<>"" then%><%= replace(oipchuldetail.FItemList(i).FDtComment," ","<br>") %><% end if %></td>
	</tr>
	</form>
	<% next %>

</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<form name="frmArrupdate" method="post" action="realStockNoEdit_process.asp">
<input type="hidden" name="arrCheckinfo" value="">
<input type="hidden" name="arrStockNo" value="">
<input type="hidden" name="arrLimitYN" value="">
</form>
<%
set oipchuldetail= Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
