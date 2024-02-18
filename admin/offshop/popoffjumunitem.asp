<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 주문
' History : 이상구 생성
'			2017.04.12 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim page, mode, designer,suplyer,shopid,itemid
dim idx

shopid = session("ssBctBigo")
page = requestCheckVar(request("page"),10)
mode = requestCheckVar(request("mode"),32)
designer = requestCheckVar(request("designer"),32)
suplyer = requestCheckVar(request("suplyer"),32)
itemid = requestCheckVar(request("itemid"),10)
idx = requestCheckVar(request("idx"),10)

if suplyer<>"10x10" then designer = suplyer
if page="" then page=1
if mode="" then mode="bybrand"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectDesigner = designer

if suplyer="10x10" then
	ioffitem.FRectShopid = shopid
	ioffitem.FRectDesignerjungsangubun = "'2','4','5'"
else
	ioffitem.FRectShopid = shopid
	ioffitem.FRectDesignerjungsangubun = "'6','8'"
end if

if (itemid<>"") then
	ioffitem.FRectDesigner =  ""
	ioffitem.FRectItemid = itemid
	ioffitem.GetOffLineJumunByItemID
elseif ((mode="bybrand") and (designer<>"")) then
	ioffitem.GetOffLineJumunItem
elseif (mode="byonbest") then
	ioffitem.FRectOrder = "byonbest"
	ioffitem.GetOnlineBestItem
elseif (mode="byonfav") then
	ioffitem.FRectOrder = "byonfav"
	ioffitem.GetOnlineBestItem
elseif (mode="byoffbest") then
	ioffitem.GetOffLineBestItem
elseif (mode="byrecent") then
	ioffitem.FRectOrder = "byrecent"
	ioffitem.GetOffLineJumunItem
elseif (mode="byetc") then
	ioffitem.FRectDesignerjungsangubun=""
	ioffitem.FRectOrder = "byetc"
	ioffitem.GetOffLineJumunItem
end if

dim i, shopsuplycash, buycash
%>
<script type='text/javascript'>

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function research(page){
	frm.page.value = page;
	frm.submit();
}

function search(frm){
	if ((frm.mode[0].checked)&&(frm.designer.value.length<1)){
		if (frm.itemid.value.length<1){
			alert('브랜드를 선택 하세요.');
			frm.designer.focus();
			return;
		}
	}

	frm.submit();
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
			}
		}
	}


	opener.ReActItems('<%= idx %>',upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value);

}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

</script>
<table width="840" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="suplyer" value="<%= suplyer %>" >
	<input type="hidden" name="idx" value="<%= idx %>" >
	<input type="hidden" name="page" value="" >
	<tr>
		<td class="a" >
			검색타입 :
			<input type="radio" name="mode" value="bybrand" <% if mode="bybrand" then response.write "checked" %> >브랜드
			<input type="radio" name="mode" value="byonbest" <% if mode="byonbest" then response.write "checked" %> >온라인 베스트
			<input type="radio" name="mode" value="byonfav" <% if mode="byonfav" then response.write "checked" %> >온라인 인기상품
			<input type="radio" name="mode" value="byoffbest" <% if mode="byoffbest" then response.write "checked" %> >오프라인 베스트
			<input type="radio" name="mode" value="byrecent" <% if mode="byrecent" then response.write "checked" %> >신 상품
			<input type="radio" name="mode" value="byetc" <% if mode="byetc" then response.write "checked" %> >기타소모품

			<br>
			브랜드 :<% drawSelectBoxOffjumunDesigner "designer",designer,shopid,suplyer %>
			<br>
			상품코드로검색 : <input type="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
		</td>
		<td class="a" align="right">
			<a href="javascript:search(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="840" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<% if ioffitem.FresultCount>0 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="right">총건수: <%= ioffitem.FTotalCount %> &nbsp; <%= Page %>/<%= ioffitem.FTotalPage %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="right"><input type="button" value="선택 아이템 추가" onclick="AddArr()"></td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF">
		<td width="20"><input type="checkbox" name="ckall" onClick="AnSelectAllFrame(this)"></td>
		<td width="50">이미지</td>
		<td width="50">브랜드ID</td>
		<td width="80">BarCode</td>
		<td width="100">상품명</td>
		<td width="80">옵션명</td>
		<td width="60">판매가</td>
		<td width="60">공급가</td>
		<td width="48">공급마진</td>
		<td width="50">갯수</td>
		<td width="70">비고</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>
	<%
	if session("ssBctDiv")="502" or session("ssBctDiv")="503" then
		shopsuplycash = ioffitem.FItemList(i).GetFranchiseSuplycash
		buycash		  = ioffitem.FItemList(i).GetFranchiseBuycash
	else
		shopsuplycash = ioffitem.FItemList(i).GetOfflineSuplycash
		buycash		  = ioffitem.FItemList(i).GetFranchiseBuycash
	end if
	%>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
	<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
	<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
	<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
	<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
	<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
	<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
	<input type="hidden" name="suplycash" value="<%= shopsuplycash %>">
	<input type="hidden" name="buycash" value="<%= buycash %>">
	<tr bgcolor="#FFFFFF">
		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td ><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<td ><%= ioffitem.FItemList(i).FMakerid %></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><%= ioffitem.FItemList(i).GetBarCode %></font></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><%= ioffitem.FItemList(i).FShopItemName %></font></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><%= ioffitem.FItemList(i).FShopItemOptionName %></font></td>
		<td align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
		<td align=right><%= FormatNumber(shopsuplycash,0) %></td>
		<td align=center>
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(shopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
		</td>
		<td ><input type="text" name="itemno" value="0" size="4" maxlength="4" onKeyDown="CheckThis(frmBuyPrc_<%= i %>);"></td>
		<td >
		<% if ioffitem.FItemList(i).Foptusing="N" then %>
		<font color="red">옵션X</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).IsSoldOut then %>
		<font color="red">판매중지</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Flimityn="Y" then %>
		<font color="blue">한정(<%= ioffitem.FItemList(i).getLimitNo %>)</font>
		<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="javascript:research('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:research('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="javascript:research('<%= i %>');">[next]</a>
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

</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->