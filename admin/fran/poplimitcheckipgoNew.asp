<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 입출고 한정 처리
' History : 2016.05.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%
dim maylimitno, mayreallimtno, currlimitea, i, addmayno

// 97% -> 98%, 2014-07-25
const C_LIMIT_PERCENT = 0.98

dim idx, alinkcode
	idx = getNumeric(requestcheckvar(request("idx"),10))
	alinkcode = requestcheckvar(request("alinkcode"),8)

if idx="" and alinkcode="" then
	response.write "구분자[1]가 없습니다."
	dbget.close() : response.end
end if

dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
	oipchul.FRectId = idx

	if idx<>"" then
		oipchul.GetIpChulMaster

		if oipchul.ftotalcount>0 then
			alinkcode = oipchul.FOneItem.Fcode
		end if
	end if

if alinkcode="" then
	response.write "구분자[2]가 없습니다."
	dbget.close() : response.end
end if

set oipchuldetail = new CIpChulStorage
	oipchuldetail.FRectStoragecode = alinkcode
	oipchuldetail.GetshopIpChulDetailCheck

%>
<script type="text/javascript">

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function ReLimitSell(itemid,itemoption,maylimitno,mayreallimtno,currlimitno){
	if (confirm('입고시 예상재고: ' + maylimitno + '\n\n한정수량 [   ' + mayreallimtno + '    ] \n\n로 조정하시겠습니까?\n\n(현재:' + currlimitno + '  추가:'+(mayreallimtno-currlimitno) + ')')){
		var popwin = window.open('/admin/newstorage/ipgoitemlimitcheckNew_process.asp?mode=addmaylimit&itemid=' + itemid + '&itemoption=' + itemoption + '&mayreallimtno=' + mayreallimtno,'doipgoitemlimitcheck','width=100 height=100');
	}
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left"></td>
	<td align="right">
        <input type="button" class="button" value="선택 한정조정" onClick="fnEditRealStockNo();">
        &nbsp;&nbsp;
		<input type="button" class="button" value="새로고침" onClick="document.location.reload();">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oipchuldetail.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oipchuldetail.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
<%
for i=0 to oipchuldetail.FresultCount-1

maylimitno = oipchuldetail.FItemList(i).FMaystockno

'/이거 안쓰기로함. 막음
'mayreallimtno = Fix(maylimitno*C_LIMIT_PERCENT)
mayreallimtno = Fix(maylimitno)

currlimitea = oipchuldetail.FItemList(i).getOptionLimitEa
%>
<form name="frmBuyPrc_<%= oipchuldetail.FItemList(i).Fiitemgubun %>_<%= oipchuldetail.FItemList(i).FItemId %>_<%= oipchuldetail.FItemList(i).FItemOption %>" method="post" >
<input type="hidden" name="mayreallimtno" value="<%=mayreallimtno%>">
<tr align='center' bgcolor="<%= oipchuldetail.FItemList(i).GetMayCheckColor %>">
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
		<font color="<%= mwdivColor(oipchuldetail.FItemList(i).FOnlineMwdiv) %>">
			<%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= oipchuldetail.FItemList(i).FItemId %>-<%= oipchuldetail.FItemList(i).FItemOption %>
		</font>
	</td>
	<td align="left">
	    <a href="javascript:PopItemSellEdit('<%= oipchuldetail.FItemList(i).FItemID %>');"><%= oipchuldetail.FItemList(i).Fiitemname %></a>
	</td>
	<td <%= chkIIF(oipchuldetail.FItemList(i).FItemOption<>"0000" and oipchuldetail.FItemList(i).FOptUsing="N","bgcolor='#FF3333'","") %>>
	    <%= oipchuldetail.FItemList(i).Fiitemoptionname %>
	</td>
	<td><%= fncolor(oipchuldetail.FItemList(i).FDanJongYn,"dj") %></td>
	<td align='right'><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
	<td><%= oipchuldetail.FItemList(i).Fbaljuitemno %></td>
	<td><%= oipchuldetail.FItemList(i).Fitemno %></td>
	<td>
	    <% if (oipchuldetail.FItemList(i).FIsNewItem = "Y") then %>
	    	<%= oipchuldetail.FItemList(i).FMaystockno %>
	    <% else %>
			<%= oipchuldetail.FItemList(i).FMaystockno %>
	    <% end if %>
	</td>
	<td><font color="<%= oipchuldetail.FItemList(i).GetIsSlodOutColor %>"><%= oipchuldetail.FItemList(i).GetIsSlodOutText %></font></td>
	<td><font color="<%= oipchuldetail.FItemList(i).GetSellYnColor %>"><%= oipchuldetail.FItemList(i).Fsellyn %></font></td>

	<% if oipchuldetail.FItemList(i).Flimityn="Y" then %>
		<td bgcolor="#FFDDDD">
			<font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font>
		</td>
	<% else %>
		<td>
			<font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font>
		</td>
	<% end if %>

	<td>
		<% if oipchuldetail.FItemList(i).FLimitYn="Y" then %>
			<%= oipchuldetail.FItemList(i).getOptionLimitEa %>
		<% end if %>
	</td>
	<td>
		<% if (oipchuldetail.FItemList(i).FLimitYn="Y") then %>
		    <% if mayreallimtno<>0 then %>
	    		<% if ((mayreallimtno<>currlimitea) or (currlimitea>mayreallimtno)) then %>
	    		<input type="button" class="button" value="->" onclick="ReLimitSell('<%= oipchuldetail.FItemList(i).FItemID %>','<%= oipchuldetail.FItemList(i).FItemOption %>','<%= maylimitno %>','<%= mayreallimtno %>','<%= currlimitea %>');">
	    		<% end if %>
		    <% end if %>
		<% end if %>
	</td>
	<td>
		<% if oipchuldetail.FItemList(i).FDtComment<>"" then%>
			<%= replace(oipchuldetail.FItemList(i).FDtComment," ","<br>") %>
		<% end if %>
	</td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<form name="frmArrupdate" method="post" action="/admin/newstorage/realStockNoEdit_process.asp">
<input type="hidden" name="arrCheckinfo" value="">
<input type="hidden" name="arrStockNo" value="">
<input type="hidden" name="arrLimitYN" value="">
</form>

<%
set oipchul = Nothing
set oipchuldetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
