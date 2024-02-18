<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 재고
' History : 이상구 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim yyyy1,mm1,dd1
dim hh1,nn1,ss1
dim makerid
dim shopid
dim idx
dim onlyusing, availstock, research
yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
dd1 = requestCheckVar(request("dd1"),2)
hh1 = requestCheckVar(request("hh1"),2)
nn1 = requestCheckVar(request("nn1"),2)
ss1 = requestCheckVar(request("ss1"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2,Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2,Cstr(day(now())))

if (hh1="") then hh1 = "00"
if (nn1="") then nn1 = "00"
if (ss1="") then ss1 = "00"

idx = requestCheckVar(request("idx"),10)
makerid = requestCheckVar(request("makerid"),32)
shopid = requestCheckVar(request("shopid"),32)
onlyusing = requestCheckVar(request("onlyusing"),2)
availstock = requestCheckVar(request("availstock"),2)
research = requestCheckVar(request("research"),2)

if (research="") and (availstock="") then availstock="on"
if (research="") and (onlyusing="") then onlyusing="on"

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid
offstock.FRecAvailStock = availstock
offstock.FRecOnlyusing = onlyusing

if idx<>"" then
	offstock.FRectIdx = idx
	offstock.GetOneJeagoMaster

	shopid = offstock.FOneItem.FShopid
	makerid = offstock.FOneItem.FMakerid

	offstock.FRectShopID = shopid
	offstock.FRectMakerid = makerid
	offstock.GetDailyStockByInputIdx
else
	offstock.GetDailyStock
end if

dim i, iptot,retot,selltot,currtot
%>

<script type='text/javascript'>

function searchItems(frm){
	if (frm.shopid.value.length<1){
		alert('샾ID를 선택하세요.');
		return;
	}

	if (frm.makerid.value.length<1){
		alert('업체ID를 선택하세요.');
		return;
	}
	frm.submit();
}

function ArrSave(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	var ret;

	upfrm.itemgubunarr.value = "";
	upfrm.shopitemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.realjeagoarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (!IsDigit(frm.realjaego.value)){
				alert('재고는 숫자만 가능합니다.');
				frm.realjaego.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.shopitemarr.value = upfrm.shopitemarr.value + frm.shopitemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.realjeagoarr.value = upfrm.realjeagoarr.value + frm.realjaego.value + "|";
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.submit();
	}
}
</script>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frm1" method="post" action="realjaegoinput.asp">
<input type="hidden" name="research" value="on">

<tr>
	<td bgcolor="#DDDDFF" width="100">IDx</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.FIdx %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">오프샾ID</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"><% drawSelectBoxOffShop "shopid",shopid %></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fshopid %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">업체ID</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "makerid",makerid  %></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fmakerid %></td>
	<% end if %>
</tr>
<tr>
	<td bgcolor="#DDDDFF">실사재고일시</td>
	<% if (idx="") then %>
	<td bgcolor="#FFFFFF"></td>
	<% else %>
	<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fjeagodate %></td>
	<% end if %>
</tr>
</table>
<br>
<% if (idx="") then %>
	<table width="800" cellspacing="1" class="a" >
	<tr >
		<td><% DrawOneDateBox yyyy1,mm1,dd1 %>
		&nbsp;
		<input type="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>시
		<input type="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>분
		<input type="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>초
		까지의 예상재고</td>
		<td align="right">
		<input type=checkbox name="availstock" <% if availstock="on" then response.write "checked" %> >유효재고만검색
		&nbsp;
		<input type=checkbox name="onlyusing" <% if onlyusing="on" then response.write "checked" %> >사용상품만검색
		<a href="javascript:searchItems(frm1);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</table>
</form>
<% else %>
</form>
<% end if %>

<% if (idx<>"") or ((shopid<>"") and (makerid<>"")) then %>
			<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
			<tr>
			<td colspan="12" align="right" bgcolor="#FFFFFF">
				실사재고 값을 수정 하신 후 "실사재고 저장" 버튼을 누르시면 값이 저장됩니다.
			</td>
			</tr>
			<tr bgcolor="#DDDDFF">
				<td width="50">이미지</td>
				<td width="86">바코드</td>
				<td width="100">상품명</td>
				<td width="80">옵션명</td>
				<td width="80">이전<br>실사일</td>
				<td width="50">이전<br>실사</td>
				<td width="50">입고</td>
				<td width="50">반품</td>
				<td width="50">판매량</td>
				<td width="50">예상재고</td>
				<td width="50">실사재고</td>
			</tr>
			<% for i=0 to offstock.FresultCount-1 %>
			<%
				iptot = iptot + offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno
				retot = retot + offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno
				selltot = selltot + offstock.FItemList(i).Fsellno
				currtot = currtot + offstock.FItemList(i).Fcurrno
			%>
			<form name="frmBuyPrc_1" >
			<input type="hidden" name="itemgubun" value="<%= offstock.FItemList(i).FItemGubun %>">
			<input type="hidden" name="shopitemid" value="<%= offstock.FItemList(i).FItemId %>">
			<input type="hidden" name="itemoption" value="<%= offstock.FItemList(i).FItemOption %>">
			<tr bgcolor="#FFFFFF">
				<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
				<td><%= offstock.FItemList(i).GetBarCode %></td>
				<td><%= offstock.FItemList(i).FItemName %></td>
				<td><%= offstock.FItemList(i).FItemOptionName %></td>
				<td align="center"><%= offstock.FItemList(i).Flastrealdate %></td>
				<td align="center"><%= offstock.FItemList(i).Flastrealno %></td>
				<td align="center"><%= offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno %></td>
				<td align="center"><%= offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno %></td>
				<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
				<% if offstock.FItemList(i).Fcurrno<1 then %>
				<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
				<% else %>
				<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
				<% end if %>

				<% if idx<>"" then %>
				<td><input type="text" name="realjaego" value="<%= offstock.FItemList(i).FinputedRealStock %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
				<% else %>
				<td><input type="text" name="realjaego" value="<%= offstock.FItemList(i).Fcurrno %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
				<% end if %>
			</tr>
			</form>
			<% next %>
			<tr bgcolor="#FFFFFF">
				<td colspan="5">total</td>
				<td align="center"></td>
				<td align="center"><%= iptot %></td>
				<td align="center"><%= retot %></td>
				<td align="center"><%= selltot %></td>
				<td align="center"><%= currtot %></td>
				<td align="center"></td>
			</tr>
			</table>
	<br>
	<table width="800" cellspacing="1" class="a" >
	<form name="frmArrupdate" method="post" action="shoprealjeago_process.asp">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="designer" value="<%= makerid %>">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="shopitemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="realjeagoarr" value="">

	<tr>
		<td align="right">재고파악 일시(정확히 입력) : <% DrawOneDateBox yyyy1,mm1,dd1 %>
		<input type="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>시
		<input type="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>분
		<input type="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>초
		<% if idx<>"" then %>
		<input type="button" value="실사재고 수정" onclick="ArrSave()">
		<% else %>
		<input type="button" value="실사재고 저장" onclick="ArrSave()">
		<% end if %>
		</td>
	</tr>
	</form>
	</table>
<% end if %>
<%
set offstock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->