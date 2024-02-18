<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim shopid, designer, shopdiv

shopid   = RequestCheckVar(request("shopid"),32)
designer = RequestCheckVar(request("designer"),32)
shopdiv  = RequestCheckVar(request("shopdiv"),32)

dim osheet
set osheet = new COrderSheet
osheet.FCurrPage = 1
osheet.Fpagesize= 1000

osheet.FRectBaljuid = shopid
osheet.FRectMakerid = designer
'샆주문중 주문접수인 것만 표시
osheet.FRectStatecd = "0"
osheet.FRectDivCodeArr = "('501','502','503')"

if (shopdiv="minus") then
    osheet.FRectMinusOnly = "on"
elseif (shopdiv="reorder") then
    osheet.FRectReOrderOnly = "on"
else
    osheet.FRectShopDiv = shopdiv
end if

if designer<>"" then
	osheet.GetOrderSheetListByBrand
else
	osheet.GetOrderSheetList
end if

dim i
%>
<script>

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function MakeBalbu(){
	var upfrm = document.frmupdate;
	var frm = document.frm;
	var pass = false;

        upfrm.masteridxarr.value = "";
	for (var i=0; i < frm.length; i++){
		if ((frm[i].name == "ck_all") && (frm[i].checked == true)) {
		        upfrm.masteridxarr.value = upfrm.masteridxarr.value + frm[i].value + "|";
		        pass = true;
                }
	}

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	if (confirm('발주서를 작성하시겠습니까?')){
		upfrm.mode.value = "makebalju";
		upfrm.submit();
	}
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmsearch" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			&nbsp;
			브랜드포함 : <% drawSelectBoxDesignerwithName "designer", designer %>
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frmsearch.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			주문구분
			<select class="select" name="shopdiv">
	     	<option value=''>전체</option>
	     	<option value='j' <%= ChkIIF(shopdiv="j","selected","") %> >직영점주문</option>
	     	<option value='f' <%= ChkIIF(shopdiv="f","selected","") %> >가맹점주문</option>
	     	<option value='f87' <%= ChkIIF(shopdiv="f87","selected","") %> >해외주문</option>
	     	<option value='reorder' <%= ChkIIF(shopdiv="reorder","selected","") %> >재주문건</option>
	     	<option value='minus' <%= ChkIIF(shopdiv="minus","selected","") %> >마이너스주문건</option>
	     	</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">

		</td>
		<td align="right">
			<input type="button" class="button" value="선택 주문 발주서 작성" onclick="MakeBalbu();" disabled>
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"></td>
		<td width="70">주문코드</td>
		<td width="100">공급처</td>
		<td width="100">주문처</td>
		<td width="70">주문상태</td>
		<td width="70">주문일</td>
		<td width="70">입고요청일</td>
		<td width="70">총주문액<br>(소비자가)</td>
		<td>포함브랜드</td>
		<td width="50">내역서</td>
	</tr>
	<form name="frm">
	<% for i=0 to osheet.FResultcount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><input type="checkbox" name="ck_all" value="<%= osheet.FItemList(i).Fidx %>" onClick="AnCheckClick(this);"></td>
		<td><a href="jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>" target="_blank"><%= osheet.FItemList(i).Fbaljucode %></a></td>
	  	<td><b><%= osheet.FItemList(i).Ftargetid %></b><br><%= osheet.FItemList(i).Ftargetname %></td>
	  	<td><%= osheet.FItemList(i).Fbaljuid %><br><%= osheet.FItemList(i).Fbaljuname %></td>
	  	<td><font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font></td>
	  	<td><%= Left(osheet.FItemList(i).FRegdate,10) %></td>
	  	<td><font color="#777777"><%= Left(osheet.FItemList(i).Fscheduledate,10) %></font></td>
	  	<td align="right"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
		<td align="left"><font color="#777777"><%= DdotFormat(osheet.FItemList(i).Fbrandlist,40) %></font></td>
	  	<td><a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>');"><img src="/images/iexplorer.gif" width=21 border=0></a></td>
	</tr>
	<% next %>
	</form>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>

</table>



<form name="frmupdate" method=post action="baljumakeroffline_process.asp">
<input type=hidden name="mode" value="makebalju">
<input type=hidden name="masteridxarr" value="">
</form>
<%

set osheet = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->