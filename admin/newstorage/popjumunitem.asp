<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim page, mode, suplyer,shopid,itemid, research
dim nubeasong, nuitem, nuitemoption
dim onoffgubun, idx

shopid = request("shopid")
page = request("page")
mode = request("mode")
suplyer = request("suplyer")
itemid = request("itemid")
research = request("research")
nubeasong = request("nubeasong")
nuitem = request("nuitem")
nuitemoption = request("nuitemoption")
onoffgubun = request("onoffgubun")
idx = request("idx")

response.write "<script>location.replace('popjumunitemNew.asp?suplyer=" + suplyer + "&changesuplyer=" + request("changesuplyer") + "&shopid=" + shopid + "&idx=" + idx + "')</script>"
dbget.close()	:	response.End

if (research<>"on") and (nubeasong="") then
	nubeasong = "on"
end if

if (research<>"on") and (nuitem="") then
	nuitem = "on"
end if

if (research<>"on") and (onoffgubun="") then
	onoffgubun = "online"
end if

if page="" then page=1
if mode="" then mode="bybrand"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = suplyer
ioffitem.FRectNoSearchUpcheBeasong = nubeasong
ioffitem.FRectNoSearchNotusingItem = nuitem
ioffitem.FRectNoSearchNotusingItemOption = nuitemoption
ioffitem.FRectItemid = itemid
if onoffgubun="offline" then
	ioffitem.GetOffShopItemList
else
	if (suplyer="") and (itemid="") then

	else
		ioffitem.GetOnLineJumunByBrand
	end if
end if

dim i, shopsuplycash, buycash
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function search(frm){
	/*
	if ((frm.suplyer.value.length<1)){
		if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}
	*/

	frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.itemno.value)){
					alert('갯수는 정수만 가능합니다.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('수량을 입력하세요.');
					frm.itemno.focus();
					return;
				}

				upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
				upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
				upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
				upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
				upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
				upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
				upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
				upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
				upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
				upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				upfrm.mwdivarr.value = upfrm.mwdivarr.value + frm.mwdiv.value + "|";

			}
		}
	}


	opener.ReActItems('<%= idx %>', upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.mwdivarr.value);

	//초기화
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				frm.cksel.checked = false;
				frm.itemno.value="0"


			}
		}
	}

}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="idx" value="<%= idx %>">
	<input type="hidden" name="page" value="<%= page %>">
	<% if (request("changesuplyer") <> "Y") then %>
	<input type="hidden" name="suplyer" value="<%= suplyer %>" >
	<% else %>
	<input type="hidden" name="changesuplyer" value="Y" >
	<% end if %>
	<input type="hidden" name="shopid" value="<%= shopid %>" >

	<tr>
		<td class="a" >
			<% if (request("changesuplyer") <> "Y") then %>
			브랜드 : <%= suplyer %>
			<% else %>
			브랜드 : <% drawSelectBoxDesignerwithName "suplyer", suplyer %>
			<% end if %>
			<input type=checkbox name="nubeasong" <% if nubeasong="on" then response.write "checked" %> >업체배송검색안함
			<input type=checkbox name="nuitem" <% if nuitem="on" then response.write "checked" %> >사용상품만
			<input type=checkbox name="nuitemoption" <% if nuitemoption="on" then response.write "checked" %> >사용옵션만

			<br>
			OnOff구분 : <input type="radio" name="onoffgubun" value="online" <% if onoffgubun="online" then response.write "checked" %> >온라인
			<input type="radio" name="onoffgubun" value="offline" <% if onoffgubun="offline" then response.write "checked" %> >오프라인
			&nbsp;
			상품코드로검색 : <input type="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
		</td>
		<td class="a" align="right">
			<a href="javascript:search(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<% if ioffitem.FresultCount>0 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="right">총건수: <%= ioffitem.FTotalCount %> &nbsp; <%= Page %>/<%= ioffitem.FTotalPage %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="right"><input type="button" value="선택 아이템 추가" onclick="AddArr()"></td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<td width="50">이미지</td>
		<td width="50">브랜드ID</td>
		<td width="80">BarCode</td>
		<td width="100">상품명</td>
		<td width="70">옵션명/<br>재고</td>
		<td width="45">판매가</td>
		<td width="45">매입가</td>
		<td width="45">매입마진</td>
		<td width="45">갯수</td>
		<td width="50">비고</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>

	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
	<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
	<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
	<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
	<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
	<input type="hidden" name="suplycash" value="<%= ioffitem.FItemList(i).Fshopsuplycash %>">
	<input type="hidden" name="buycash" value="<%= ioffitem.FItemList(i).Fshopsuplycash %>">
	<input type="hidden" name="mwdiv" value="<%= ioffitem.FItemList(i).Fmwdiv %>">

	<tr bgcolor="#FFFFFF">
		<td rowspan=2><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan=2><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= ioffitem.FItemList(i).FMakerid %></td>
		<td ><a href="javascript:PopItemSellEdit('<%= ioffitem.FItemList(i).FShopItemID %>');"><%= ioffitem.FItemList(i).GetBarCodeBoldStr %></a></td>
		<td ><%= ioffitem.FItemList(i).FShopItemName %></td>
		<td ><%= ioffitem.FItemList(i).FShopItemOptionName %></td>
		<td rowspan=2 align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
		<td rowspan=2 align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopsuplycash,0) %></td>
		<td rowspan=2 align=center>
		<font color="<%= ioffitem.FItemList(i).getMwDivColor %>"><%= ioffitem.FItemList(i).getMwDivName %></font><br>
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(ioffitem.FItemList(i).Fshopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
		</td>
		<td rowspan=2 ><input type="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
		<td rowspan=2 >

		<% if ioffitem.FItemList(i).Foptusing="N" then %>
		<font color="red">옵션x</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).IsSoldOut then %>
		<font color="red">판매중지</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Flimityn="Y" then %>
		<font color="blue">한정(<%= ioffitem.FItemList(i).getLimitNo %>)</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Fpreorderno<>0 then %>
		기주문:<%= ioffitem.FItemList(i).Fpreorderno %>
		<% end if %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=3>
			<font color="#444444">
			[<%= Left(ioffitem.FItemList(i).Flastrealdate,10) %>]
			<%= ioffitem.FItemList(i).Flastrealno %>
			+
			<%= ioffitem.FItemList(i).Fipno %>
			<% if ioffitem.FItemList(i).Fchulno<0 then %>
			-
			<% else %>
			+
			<% end if %>
			<%= Abs(ioffitem.FItemList(i).Fchulno) %>
			-
			<%= ioffitem.FItemList(i).Fsellno %>
			-
			<%= ioffitem.FItemList(i).Fipkumdiv4 %>
			-
			<%= ioffitem.FItemList(i).Fipkumdiv2 %>
			</font>
			<br>
			<%= ioffitem.FItemList(i).Fmaxsellday %>일[<%= ioffitem.FItemList(i).Fsell7days %>]
			<% if ioffitem.FItemList(i).Fmaxsellday<>0 then %>
			일평균[<%= CLng(ioffitem.FItemList(i).Fsell7days/ioffitem.FItemList(i).Fmaxsellday*10)/10 %>]
			<% else %>
			일평균[-]
			<% end if %>
			적정[<%= ioffitem.FItemList(i).Frequireno %>]
			오프[<%= ioffitem.FItemList(i).Foffjupno+ioffitem.FItemList(i).Foffconfirmno %>]

			결제[<%= ioffitem.FItemList(i).Fipkumdiv4 %>]
			무통[<%= ioffitem.FItemList(i).Fipkumdiv2 %>]

			<% if ioffitem.FItemList(i).Getshortageno<0 then %>
			부족수량[<font color="#CC1111"><b><%= ioffitem.FItemList(i).Getshortageno %></b></font>]
			<% else %>
			부족수량[+<%= ioffitem.FItemList(i).Getshortageno %>]
			<% end if %>

		</td>
		<td align=center>
			<% if ioffitem.FItemList(i).Fcurrno<1 then %>
			<font color="#CC1111"><b><%= ioffitem.FItemList(i).Fcurrno %></b></font>
			<% else %>
			<%= ioffitem.FItemList(i).Fcurrno %>
			<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ioffitem.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StarScrollPage to ioffitem.FScrollCount + ioffitem.StarScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="frmArrupdate" method="post" action="">
<input type="hidden" name="mode" value="arrins">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->