<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/base64.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp"-->
<%
''oitemcouponmaster, 
dim itemcouponidx
dim ocouponitemlist
dim page, makerid,sailyn, invalidmargin
dim sRectItemidArr, tmpArr
dim dispCate, couponGubun, openstate
dim duppexists, duppnvup
dim exceptnvcpn

itemcouponidx   = requestCheckvar(request("itemcouponidx"),10)
makerid         = requestCheckvar(request("makerid"),32)
page            = requestCheckvar(request("page"),10)
sailyn          = requestCheckvar(request("sailyn"),10)
invalidmargin   = requestCheckvar(request("invalidmargin"),10)
sRectItemidArr  = Trim(request("sRectItemidArr"))
dispCate		= requestCheckvar(request("disp"),16)
''onlynv          = requestCheckvar(request("onlynv"),10)
couponGubun		= requestCheckvar(request("couponGubun"),10)
openstate       = requestCheckvar(request("openstate"),10)
duppexists		= requestCheckvar(request("duppexists"),10)
exceptnvcpn		= requestCheckvar(request("exceptnvcpn"),10)
duppnvup		= requestCheckvar(request("duppnvup"),10)

if (duppexists="") then duppnvup=""

'상품코드 검색 분해/재조립
if sRectItemidArr<>"" then
	sRectItemidArr = Replace(sRectItemidArr," ",",")
	sRectItemidArr = Replace(sRectItemidArr,vbCrLf,",")
	tmpArr = split(sRectItemidArr,",")
	sRectItemidArr = ""
	for i=0 to uBound(tmpArr)
		if isNumeric(tmpArr(i)) then
			sRectItemidArr = sRectItemidArr & chkIIF(sRectItemidArr<>"",",","") & trim(tmpArr(i))
		end if
	next
end if

itemcouponidx=trim(itemcouponidx)
if itemcouponidx="" then itemcouponidx=0
if not isNumeric(itemcouponidx) then itemcouponidx=0
if page="" then page=1


'set oitemcouponmaster = new CItemCouponMaster
'oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
'oitemcouponmaster.GetOneItemCouponMaster
'rw openstate

set ocouponitemlist = new CItemCouponMaster
ocouponitemlist.FPageSize=50
ocouponitemlist.FCurrPage=page
ocouponitemlist.FRectItemCouponIdx = itemcouponidx
ocouponitemlist.FRectMakerid = makerid
ocouponitemlist.FRectSailYn = sailyn
ocouponitemlist.FRectInvalidMargin = invalidmargin
ocouponitemlist.FRectsRectItemidArr = sRectItemidArr
ocouponitemlist.FRectDispCate		= dispCate
'ocouponitemlist.FRectOnlyValid      = "on"
ocouponitemlist.FRectOnlyValid      = "" '// 종료된 쿠폰도 검색 가능하도록 MD에서 요청
ocouponitemlist.FRectCouponGubun    = couponGubun ''CHKIIF(onlynv<>"","V","")
ocouponitemlist.FRectOpenState = openstate
ocouponitemlist.FRectDuppexists = duppexists
ocouponitemlist.FRectDuppNvUpCase = duppnvup
ocouponitemlist.FRectExceptnvcpn = exceptnvcpn

if ocouponitemlist.FRectInvalidMargin="Y" then  ''속도개선/ noPaging
	ocouponitemlist.FPageSize = 500
end if

ocouponitemlist.GetItemCouponItemListMulti

dim i


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

//function AddIttems(){
//	frmbuf.submit();
//}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

function EditArr(){
	var upfrm = document.frmbuf;
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

	upfrm.itemidarr.value = "";
	upfrm.couponbuypricearr.value = "";
    upfrm.couponsellcasharr.value = "";
    
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsDigit(frm.couponbuyprice.value)){
					alert('매입가는 숫자만 가능합니다.');
					frm.couponbuyprice.focus();
					return;
				}

				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
				upfrm.couponbuypricearr.value = upfrm.couponbuypricearr.value + frm.couponbuyprice.value + ",";
				upfrm.couponsellcasharr.value = upfrm.couponsellcasharr.value + frm.couponsellcash.value + ",";
                
			}
		}
	}



	if (confirm('선택 상품을 수정 하시겠습니까?')){
		frmbuf.mode.value="modicouponitemMulti"
//		frmbuf.submit();
	}
}

function DelArr(){
	var upfrm = document.frmbuf;
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.itemidarr.value = "";
	upfrm.couponbuypricearr.value = "";
	upfrm.itemcouponidxarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
				upfrm.itemcouponidxarr.value = upfrm.itemcouponidxarr.value + frm.itemcouponidx.value + ",";
			}
		}
	}


	if (confirm('선택 상품을 삭제 하시겠습니까?')){
		upfrm.mode.value="delcouponitemmulti"
		frmbuf.submit();
	}
}

function couponCodeClickSearch(cc) {
	$("#itemcouponidx").empty().val(cc);
	document.frm.submit();
}
</script>


<form name="frm" method="POST" action="/admin/shopmaster/itemcouponitemlisteidtMulti.asp">
<input type="hidden" name="page" value="1">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a">
<tr height="25" bgcolor="F4F4F4">
    <td valign="middle" bgcolor="F4F4F4" colspan="3">
    	쿠폰코드 : <input type="text" name="itemcouponidx" id="itemcouponidx" value="<%= CHKIIF(itemcouponidx="0","",itemcouponidx) %>" size="8">
    	&nbsp;&nbsp;
    	브랜드 : <% drawSelectBoxDesignerWithName "makerid",makerid %>
    	&nbsp;&nbsp;
    	상태 : 
        <select name="openstate">
            <option value="" <%=CHKIIF(openstate="","selected","")%> >오픈,발급대기,발급예약
            <option value="7" <%=CHKIIF(openstate="7","selected","")%> >오픈
            <option value="0,6" <%=CHKIIF(openstate="0,6","selected","")%> >발급대기, 발급예약
            <option value="0" <%=CHKIIF(openstate="0","selected","")%> >발급대기
            <option value="6" <%=CHKIIF(openstate="6","selected","")%> >발급예약
			<option value="9" <%=CHKIIF(openstate="9","selected","")%> >종료
        </select>
		&nbsp;&nbsp;
		쿠폰구분 
		<select name="couponGubun">
		    <option value="" <%=CHKIIF(couponGubun="","selected","") %> >전체
		    <option value="C" <%=CHKIIF(couponGubun="C","selected","") %> >일반
		    <option value="V" <%=CHKIIF(couponGubun="V","selected","") %> >네이버전용쿠폰
		    <option value="P" <%=CHKIIF(couponGubun="P","selected","") %> >지정인발급
		    <option value="T" <%=CHKIIF(couponGubun="T","selected","") %> >타겟(E-mail특가)
		</select>
	</td>
</tr>
<tr height="25" bgcolor="F4F4F4">
    <td valign="middle" bgcolor="F4F4F4" colspan="3">
	전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
<tr>
	<td bgcolor="F4F4F4">
    	<input type="checkbox" name="sailyn" value="Y" <% if sailyn="Y" then response.write "checked" %> >세일중인 상품 검색<br>
        <input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >마진부족(or 역마진) 상품 검색<br>
		<input type="checkbox" name="exceptnvcpn" value="Y" <% if exceptnvcpn="Y" then response.write "checked" %> >네어버쿠폰 제외 상품(브랜드) 검색<br><br>

		<% if (FALSE) then %>
		<input type="checkbox" name="onlynv" value="Y" <% if onlynv="Y" then response.write "checked" %> >네어버쿠폰만 검색<br>
		<% end if %>
		<br>
        <input type="checkbox" name="duppexists" value="Y" <% if duppexists="Y" then response.write "checked" %> >중복쿠폰검색
		&nbsp;
		(<input type="checkbox" name="duppnvup" value="Y" <% if duppnvup="Y" then response.write "checked" %> >(NV쿠폰적용가>일반쿠폰적용가))
		<br>
		* 네이버쿠폰과 일반쿠폰 중복 안되는케이스 <br>
		: 일반무료배송쿠폰-네이버쿠폰 적용불가 <br>
		: 일반쿠폰할인액>네이버쿠폰할인액 (이런경우 일반쿠폰으로 진행)

    </td>
    <td valign="middle" bgcolor="F4F4F4">
        상품코드:<br><textarea name="sRectItemidArr" style="width:350px; height:50px;"><%= sRectItemidArr %></textarea>
	</td>
    <td valign="middle" align="right" bgcolor="F4F4F4" rowspan="2" >
        <input type="button" class="button" value="등록된 상품 검색" onclick="document.frm.submit();" style="height:40px;">
    </td>
</tr>
</table>
</form>

<span>* <font color="red">쿠폰적용시 매입가 0</font>인 경우는 현재의 매입가로 설정됩니다. (매입가 조정이 없는경우는 0으로 설정할것!)</span>
<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
	<td colspan="14" align="left">
	<% if (FALSE) then %><input type="button" class="button" value="선택상품수정" onclick="EditArr();"><% end if %>
	<input type="button" class="button" value="선택상품삭제" onclick="DelArr();">
	</td>
	<td colspan="3" align="right"><%=FormatNumber(ocouponitemlist.FTotalCount,0) %> 건</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="20"><input type="checkbox" name="ckall" onclick="AnSelectAllFrame(this.checked)"></td>
	<td width="50">쿠폰코드</td>
	<td width="80">쿠폰구분</td>
	<td width="50">쿠폰상태</td>
	<td width="80">할인금액</td>
	<td width="50">이미지</td>
	<td width="80">브랜드</td>
	<td width="60">상품번호</td>
	<td >상품명</td>
	<td width="60">판매<br>상태</td>
	<td width="60">현재 판매가</td>
	<td width="60">현재 매입가</td>
	<td width="40">매입<br>구분</td>
	<td width="50">현재 마진</td>
	<td width="60">쿠폰적용시<br>판매가</td>
	<td width="60">쿠폰적용시<br>매입가</td>
	<td width="60">쿠폰적용시<br>마진(현재가 비교)</td>
	<!-- <td width="60">쿠폰적용시<br>마진(등록시)</td> -->
</tr>
<% for i=0 to ocouponitemlist.FResultCount - 1 %>
<form name="frmBuyPrc_<%= ocouponitemlist.FitemList(i).FItemID %>" method="post" onSubmit="return false;" action="do_itemcoupon.asp">
<input type="hidden" name="itemid" value="<%= ocouponitemlist.FitemList(i).FItemID %>">
<tr bgcolor="#FFFFFF">
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><a href="#" onclick="couponCodeClickSearch('<%=ocouponitemlist.FitemList(i).Fitemcouponidx%>');"><%= ocouponitemlist.FitemList(i).Fitemcouponidx %></a><input type="hidden" name="itemcouponidx" value="<%= ocouponitemlist.FitemList(i).Fitemcouponidx %>"></td>
	<td align="center"><font color="<%= ocouponitemlist.FItemList(i).getMasterCouponGubunColor %>"><%= ocouponitemlist.FItemList(i).getMasterCouponGubunName %></font></td>
	<td align="center"><font color="<%= ocouponitemlist.FItemList(i).GetMasterOpenStateColor %>"><%= ocouponitemlist.FitemList(i).GetMasterOpenStateName %></font></td>
	<td><%= ocouponitemlist.FItemList(i).GetMasterDiscountStr %></td>
	<td ><img src="<%= ocouponitemlist.FitemList(i).FSmallimage %>"width="50"></td>
	<td><%= ocouponitemlist.FitemList(i).FMakerid %></td>
	<td align="center"><%= ocouponitemlist.FitemList(i).FItemID %>
    	
	</td>
	<td ><%= ocouponitemlist.FitemList(i).FItemName %></td>
	<td ><%= ocouponitemlist.FitemList(i).getItemSellStateName %></td>
	<td align="right">
	    <% if (ocouponitemlist.FitemList(i).Forgprice>ocouponitemlist.FitemList(i).FSellcash) then %>
	    <font color=#AAAAAA><%=ocouponitemlist.FitemList(i).getSaleDiscountProStr%><%= FormatNumber(ocouponitemlist.FitemList(i).Forgprice,0) %></font>
	    <br><%= FormatNumber(ocouponitemlist.FitemList(i).FSellcash,0) %>
	    <% else %>
	    <%= FormatNumber(ocouponitemlist.FitemList(i).FSellcash,0) %>
	    <% end if %>
	</td>
	<td align="right">
	    <% if (ocouponitemlist.FitemList(i).Forgprice>ocouponitemlist.FitemList(i).FSellcash) then %>
	    <font color=#AAAAAA><%= FormatNumber(ocouponitemlist.FitemList(i).Forgsuplycash,0) %></font>
	    <br><%= FormatNumber(ocouponitemlist.FitemList(i).FBuycash,0) %>
	    <% else %>
	    <%= FormatNumber(ocouponitemlist.FitemList(i).FBuycash,0) %>
	    <% end if %>
	</td>
	<td align="center"><font color="<%= ocouponitemlist.FitemList(i).GetMwDivColor %>"><%= ocouponitemlist.FitemList(i).GetMwDivName %></font>
	<td align="center">
	    <% if (ocouponitemlist.FitemList(i).Forgprice>ocouponitemlist.FitemList(i).FSellcash) then %>
	    <font color=#AAAAAA><%= FormatNumber(ocouponitemlist.FitemList(i).GetOriginMargin,0) %>%</font>
	    <br><%= ocouponitemlist.FitemList(i).GetCurrentMargin %>%
	    <% else %>
	    <%= ocouponitemlist.FitemList(i).GetCurrentMargin %>%
	    <% end if %>
	</td>
	<td align="right"><%= FormatNumber(ocouponitemlist.FitemList(i).GetCouponSellcash,0) %>
	<% if ocouponitemlist.FitemList(i).Fitemcoupontype="3" then %>
	<br><font color="red">(무료배송)</font>
	<% end if %>
	<input type="hidden" name="couponsellcash" value="<%=ocouponitemlist.FitemList(i).GetCouponSellcash%>">
	</td>
	<td align="right">
	    <input type="text" name="couponbuyprice" value="<%= ocouponitemlist.FitemList(i).Fcouponbuyprice %>" size="7" maxlength="9" style="border:1px #999999 solid; text-align=right" onKeyDown="CheckThis(this.form);">
	    <% if (ocouponitemlist.FitemList(i).getMayCouponBuyPriceByMaginType<>ocouponitemlist.FitemList(i).Fcouponbuyprice) then %>
	    <br><%=ocouponitemlist.FitemList(i).getMayCouponBuyPriceByMaginType%>
	    <% end if %>
	</td>
	<td align="center"> 
	     <font color="<%= ocouponitemlist.FitemList(i).GetCouponMarginColor %>"><%= ocouponitemlist.FitemList(i).GetCouponMargin %></font>%
    	    <% if ocouponitemlist.FitemList(i).Fitemcoupontype="3" then %>
    	    <br><font color="red"><%= ocouponitemlist.FitemList(i).GetFreeBeasongCouponMargin %></font>%
    	    <% end if %>
	</td>
	<!--
	<td align="center"> 
	    <%if not isNull(ocouponitemlist.FitemList(i).Fcouponmargin) then %>
	     <font color="<%if ocouponitemlist.FitemList(i).Fitemcoupontype="3" then%>red<%else%><%= ocouponitemlist.FitemList(i).GetCouponMarginColor %><%end if%>">
	    <%= CLNG(ocouponitemlist.FitemList(i).Fcouponmargin*100)/100 %></font>%
	    <%end if%>
	</td>
	-->
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="17" align="center">
		<% if ocouponitemlist.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ocouponitemlist.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocouponitemlist.StarScrollPage to ocouponitemlist.FScrollCount + ocouponitemlist.StarScrollPage - 1 %>
			<% if i>ocouponitemlist.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocouponitemlist.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
set ocouponitemlist = Nothing
%>
<form name="frmbuf" method="post" action="/admin/shopmaster/itemcoupon_process.asp">
<input type="hidden" name="mode" value="addcouponitemarr">
<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="couponbuypricearr" value="">
<input type="hidden" name="couponsellcasharr" value="">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="sailyn" value="<%= sailyn %>">
<input type="hidden" name="defaultmargin" value="">

<input type="hidden" name="itemcouponidxarr"  value="">
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
