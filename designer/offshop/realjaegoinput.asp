<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim yyyy1,mm1,dd1
dim hh1,nn1,ss1
dim makerid
dim shopid
dim idx
dim onlyusing, availstock, research

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
hh1 = request("hh1")
nn1 = request("nn1")
ss1 = request("ss1")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2,Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2,Cstr(day(now())))

if (hh1="") then hh1 = "00"
if (nn1="") then nn1 = "00"
if (ss1="") then ss1 = "00"

idx = request("idx")
makerid = session("ssBctID")
shopid = request("shopid")
onlyusing = request("onlyusing")
availstock = request("availstock")
research = request("research")
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
<script language='javascript'>
function searchItems(frm){
	if (frm.shopid.value.length<1){
		alert('샾을 선택하세요.');
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm1" method="post" action="realjaegoinput.asp">
	<input type=hidden name=research value="on">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width="100">IDx</td>
		<% if (idx="") then %>
		<td bgcolor="#FFFFFF"></td>
		<% else %>
		<td bgcolor="#FFFFFF"><%= offstock.FOneItem.FIdx %></td>
		<% end if %>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">OffSHOP ID</td>
		<input type=hidden name="shopid" value="<%= shopid %>">
		<% if (idx="") then %>
		<td bgcolor="#FFFFFF"><%= shopid %></td>
		<% else %>
		<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fshopid %></td>
		<% end if %>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<% if (idx="") then %>
		<td bgcolor="#FFFFFF"><%= makerid %></td>
		<% else %>
		<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fmakerid %></td>
		<% end if %>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">최종실사재고일시</td>
		<% if (idx="") then %>
		<td bgcolor="#FFFFFF"></td>
		<% else %>
		<td bgcolor="#FFFFFF"><%= offstock.FOneItem.Fjeagodate %></td>
		<% end if %>
	</tr>
</table>

<!-- 검색 시작 -->
<% if (idx="") then %>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<% DrawOneDateBox yyyy1,mm1,dd1 %>
			&nbsp;
			<input type="text" class="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>시
			<input type="text" class="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>분
			<input type="text" class="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>초
			까지의 예상재고
			&nbsp;
			<input type=checkbox name="availstock" <% if (availstock="on") then response.write "checked" %> >유효재고만검색
			&nbsp;
			<input type=checkbox name="onlyusing" <% if (onlyusing="on") then response.write "checked" %> >사용상품만검색
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:javascript:searchItems(frm1);">
		</td>
	</tr>
	</form>
</table>

<% else %>
</form>
<% end if %>

<p>

<% if (idx<>"") or ((shopid<>"") and (makerid<>"")) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			실사재고 값을 수정 하신 후 페이지 하단 [실사재고 저장] 버튼을 누르시면 값이 저장됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">이미지</td>
		<td width="86">바코드</td>
		<td>상품명</td>
		<td>옵션명</td>
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
	<tr align="center" bgcolor="#FFFFFF">
		<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
		<td><%= offstock.FItemList(i).GetBarCode %></td>
		<td align="left"><%= offstock.FItemList(i).FItemName %></td>
		<td><%= offstock.FItemList(i).FItemOptionName %></td>
		<td><%= offstock.FItemList(i).Flastrealdate %></td>
		<td><%= offstock.FItemList(i).Flastrealno %></td>
		<td><%= offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno %></td>
		<td><%= offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno %></td>
		<td><%= offstock.FItemList(i).Fsellno %></td>
		<% if offstock.FItemList(i).Fcurrno<1 then %>
		<td><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
		<% else %>
		<td><%= offstock.FItemList(i).Fcurrno %></td>
		<% end if %>

		<% if idx<>"" then %>
		<td><input type="text" class="text" name="realjaego" value="<%= offstock.FItemList(i).FinputedRealStock %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
		<% else %>
		<td><input type="text" class="text" name="realjaego" value="<%= offstock.FItemList(i).Fcurrno %>" size="4" maxlength=8 style="border:1px #999999 solid; text-align=center"></td>
		<% end if %>
	</tr>
	</form>
	<% next %>
	<tr height="25" align="center" bgcolor="#FFFFFF">
		<td colspan="5">합계</td>
		<td align="center"></td>
		<td align="center"><%= iptot %></td>
		<td align="center"><%= retot %></td>
		<td align="center"><%= selltot %></td>
		<td align="center"><%= currtot %></td>
		<td align="center"></td>
	</tr>
</table>

<p>

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
		<% if idx<>"" then %>
		<td align="right">재고파악 일시(정확히 입력) : <% DrawOneDateBox Left(offstock.FOneItem.Fjeagodate,4),mid(offstock.FOneItem.Fjeagodate,6,2),mid(offstock.FOneItem.Fjeagodate,9,2) %>
			<input type="text" class="text" name="hh1" value="<%= mid(offstock.FOneItem.Fjeagodate,12,2) %>" size=2 maxlength=2>시
			<input type="text" class="text" name="nn1" value="<%= mid(offstock.FOneItem.Fjeagodate,15,2) %>" size=2 maxlength=2>분
			<input type="text" class="text" name="ss1" value="<%= mid(offstock.FOneItem.Fjeagodate,18,2) %>" size=2 maxlength=2>초
			<% if idx<>"" then %>
			<input type="button" class="button" value="실사재고 수정" onclick="ArrSave()">
			<% else %>
			<input type="button" class="button" value="실사재고 저장" onclick="ArrSave()">
			<% end if %>
		<% else %>
			<td align="right">재고파악 일시(정확히 입력) : <% DrawOneDateBox yyyy1,mm1,dd1 %>
			<input type="text" class="text" name="hh1" value="<%= hh1 %>" size=2 maxlength=2>시
			<input type="text" class="text" name="nn1" value="<%= nn1 %>" size=2 maxlength=2>분
			<input type="text" class="text" name="ss1" value="<%= ss1 %>" size=2 maxlength=2>초
			<% if idx<>"" then %>
			<input type="button" class="button" value="실사재고 수정" onclick="ArrSave()">
			<% else %>
			<input type="button" class="button" value="실사재고 저장" onclick="ArrSave()">
			<% end if %>
		<% end if %>
		</td>
	</tr>
	</form>
	</table>
<% end if %>
<%
set offstock = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->