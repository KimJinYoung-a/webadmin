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
dim itemcouponidx
dim oitemcouponmaster, ocouponitemlist
dim page, makerid,sailyn,mwdiv, invalidmargin
dim sRectItemidArr, tmpArr
dim dispCate

itemcouponidx   = requestCheckvar(request("itemcouponidx"),10)
makerid         = requestCheckvar(request("makerid"),32)
page            = requestCheckvar(request("page"),10)
sailyn          = requestCheckvar(request("sailyn"),10)
mwdiv           = requestCheckvar(request("mwdiv"),2)
invalidmargin   = requestCheckvar(request("invalidmargin"),10)
sRectItemidArr  = Trim(request("sRectItemidArr"))
dispCate		= requestCheckvar(request("disp"),16)

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

if itemcouponidx="" then itemcouponidx=0
if page="" then page=1


set oitemcouponmaster = new CItemCouponMaster
oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
oitemcouponmaster.GetOneItemCouponMaster

set ocouponitemlist = new CItemCouponMaster
ocouponitemlist.FPageSize=50
ocouponitemlist.FCurrPage=page
ocouponitemlist.FRectItemCouponIdx = itemcouponidx
ocouponitemlist.FRectMakerid = makerid
ocouponitemlist.FRectSailYn = sailyn
ocouponitemlist.FRectMwdiv = mwdiv
ocouponitemlist.FRectInvalidMargin = invalidmargin
ocouponitemlist.FRectsRectItemidArr = sRectItemidArr
ocouponitemlist.FRectDispCate		= dispCate

if ocouponitemlist.FRectInvalidMargin="Y" then  ''속도개선/ noPaging
	ocouponitemlist.FPageSize = 500
end if
ocouponitemlist.GetItemCouponItemList

dim i


%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function AddIttems(){
	frmbuf.submit();
}

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
    var limitMarPrc = 0;
	var limitMarPer = 0;
	var resultitemid = "";
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

				if (frm.mwdiv.value!="M"){
					limitMarPrc = frm.orgsuplycash.value-((frm.orgprice.value-frm.couponsellcash.value)*0.5);
					limitMarPer = (frm.couponsellcash.value-limitMarPrc)/frm.couponsellcash.value*100;
					if(parseInt(limitMarPrc)>parseInt(frm.couponbuyprice.value)) {
						resultitemid = resultitemid + upfrm.itemidarr.value + "\n"
					}
				}
			}
		}
	}

	if (resultitemid!=""){
		if(!confirm('업체 할인 분담율이 50%를 넘는 상품이 있습니다.\n\n입력하신 내용이 정확합니까?')){;
			return;
		}
	}

	if (confirm('선택 상품을 수정 하시겠습니까?')){
		frmbuf.mode.value="modicouponitemarr"
		frmbuf.submit();
	}
}

function DelArr(){
	var upfrm = document.frmbuf;
	var frm;
	var pass = false;

	<% if oitemcouponmaster.FOneItem.FisuedCount>0 and oitemcouponmaster.FOneItem.Fopenstate=7 then %>
	if (!confirm('고객께 발급된 쿠폰이 존재합니다!\n\n계속 진행 하시겠습니까?')){
		return;
	}
	<% end if %>

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

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + ",";
			}
		}
	}


	if (confirm('선택 상품을 삭제 하시겠습니까?')){
		upfrm.mode.value="delcouponitemarr"
		frmbuf.submit();
	}
}

//브랜드 상품 일괄삭제
function DelBrandAll() {
	var upfrm = document.frmbuf;
	var makerid = upfrm.makerid.value;

	if(makerid=="") {
		alert("검색한 브랜드가 없습니다.\n브랜드 검색 후 이용해주세요.")
		return;
	}

	if (confirm(makerid+'브랜드의 상품을 모두 삭제 하시겠습니까?')){
		upfrm.mode.value="delBrandAll"
		frmbuf.submit();
	}
}

// Old
function AddNewCouponItem(targetcomp){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist3.asp?dispyn=Y&sellyn=Y&sailyn=N&target=" + targetcomp, "AddNewCouponItem", "width=800,height=500,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 새상품 추가 팝업
function addnewItem(couponCD,evtCd){
		var popwin;
		<% if (oitemcouponmaster.FOneItem.FcouponGubun="V") and (oitemcouponmaster.FOneItem.Fitemcoupontype="1") then ''네이버 전용쿠폰인경우 && %쿠폰인경우 %>
		    popwin = window.open("/admin/itemmaster/pop_itemAddInfo_NvCpn.asp?icpnIdx=<%=itemcouponidx%>&PR=V&sellyn=Y&usingyn=Y&sailyn=&defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/admin/shopmaster/itemcoupon_Process.asp?itemcouponidx=" + couponCD, "popup_item_NvCpn", "width=1000,height=600,scrollbars=yes,resizable=yes");
	    <% else %>
		if ( evtCd > 0 ){
			popwin = window.open("/admin/eventmanage/common/pop_eventitem_addinfo.asp?defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/admin/shopmaster/itemcoupon_Process.asp?itemcouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}else{
			popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&sailyn=N&defaultmargin=<%=oitemcouponmaster.FOneItem.FDefaultMargin%>&acURL=/admin/shopmaster/itemcoupon_Process.asp?itemcouponidx=" + couponCD, "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		}
	    <% end if %>
		popwin.focus();
}

// 클립보드로 복사
function fnCBCopy(iid) {
	var doc = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + iid + "&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>";
	clipboardData.setData("Text", doc);
	alert('선택하신 상품의 링크가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}

// 타겟쿠폰 링크 팝업
function fnPopLinkCopy(iid) {
	var popwin;
	popwin = window.open("/admin/shopmaster/pop_targetItemcouponView.asp?icpidx=<%=oitemcouponmaster.FOneItem.FItemCouponIdx%>&iid=" + iid, "popTagerCpLink", "width=500,height=325,scrollbars=yes,resizable=yes");
	popwin.focus();
}
</script>


<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="right"><input type="button" class="button" value="새 상품 추가 + " onclick="addnewItem('<%= itemcouponidx %>');"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100">쿠폰명</td>
	<td bgcolor="#FFFFFF"><%= oitemcouponmaster.FOneItem.Fitemcouponname %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >할인율</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetDiscountStr %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >적용기간</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.Fitemcouponstartdate %> ~ <%= oitemcouponmaster.FOneItem.Fitemcouponexpiredate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >마진구분</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetMargintypeName %> <% if oitemcouponmaster.FOneItem.FDefaultMargin<>0 then %>- (<%= oitemcouponmaster.FOneItem.FDefaultMargin %>%) <% End IF %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >발급 상태</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName & chkIIF(oitemcouponmaster.FOneItem.FisuedCount>0," / 발급쿠폰수 : <b>" & FormatNumber(oitemcouponmaster.FOneItem.FisuedCount,0) & "</b>","") %>
	</td>
</tr>
</table>

<form name="frm" method="POST" action="itemcouponitemlisteidt.asp">
<input type="hidden" name="page" value="1">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
<tr height="25" bgcolor="F4F4F4">
    <td valign="middle" bgcolor="F4F4F4">
    	<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>" >
    	브랜드 : <% drawSelectBoxDesignerWithName "makerid",makerid %><br>
    	할인여부 :
		<select name="sailyn" class="select">
		<option value="">전체</option>
		<option value="Y" <%=chkIIF(sailyn="Y","selected","")%>>할인중</option>
		<option value="N" <%=chkIIF(sailyn="N","selected","")%>>안함</option>
		</select>
		매입구분 :
		<select name="mwdiv" class="select">
		<option value="">전체</option>
		<option value="MW" <%=chkIIF(mwdiv="MW","selected","")%>>매입+위탁</option>
		<option value="M" <%=chkIIF(mwdiv="M","selected","")%>>매입</option>
		<option value="W" <%=chkIIF(mwdiv="W","selected","")%>>위탁</option>
		<option value="U" <%=chkIIF(mwdiv="U","selected","")%>>업체</option>
		</select>
		<br />
        <label><input type="checkbox" name="invalidmargin" value="Y" <% if invalidmargin="Y" then response.write "checked" %> >마진부족(or 역마진) 상품 검색</label>
    </td>
    <td valign="middle">
        상품코드:<br><textarea name="sRectItemidArr" style="width:350px; height:50px;"><%= sRectItemidArr %></textarea>
	</td>
    <td valign="middle" align="right" bgcolor="F4F4F4" rowspan="2">
        <input type="button" class="button" value="등록된 상품 검색" onclick="document.frm.submit();" style="height:40px;">
    </td>
</tr>
<tr>
	<td bgcolor="F4F4F4" style="white-space:nowrap;padding-left:5px;" colspan="2">전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
</tr>
</table>
</form>

<span>* <font color="red">쿠폰적용시 매입가 0</font>인 경우는 현재의 매입가로 설정됩니다. (매입가 조정이 없는경우는 0으로 설정할것!)</span>
<br>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="left">
	<input type="button" class="button" value="선택상품수정" onclick="EditArr();" />
	<input type="button" class="button" value="선택상품삭제" onclick="DelArr();" />
	<% if not(isNull(makerid) or makerid="") then%>
	&nbsp;/&nbsp; <input type="button" class="button" value="<%=makerid%>상품 일괄삭제" onclick="DelBrandAll();" style="background-color:#FFCCCC;" />
	<% end if %>
	</td>
	<td colspan="3" align="right"><%=FormatNumber(ocouponitemlist.FTotalCount,0) %> 건</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="20"><input type="checkbox" name="ckall" onclick="AnSelectAllFrame(this.checked)"></td>
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
	<td ><img src="<%= ocouponitemlist.FitemList(i).FSmallimage %>"width="50"></td>
	<td><%= ocouponitemlist.FitemList(i).FMakerid %></td>
	<td align="center"><%= ocouponitemlist.FitemList(i).FItemID %>
    	<% if oitemcouponmaster.FOneItem.FcouponGubun="T" then %>
    	<input type="button" class="button" value="URL생성" onClick="fnPopLinkCopy('<%=ocouponitemlist.FitemList(i).FItemID%>')">
    	<% end if %>
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
	<input type="hidden" name="orgsuplycash" value="<%=ocouponitemlist.FitemList(i).Forgsuplycash%>">
	<input type="hidden" name="orgprice" value="<%=ocouponitemlist.FitemList(i).Forgprice%>">
	<input type="hidden" name="mwdiv" value="<%=ocouponitemlist.FitemList(i).FMwDiv%>">
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
			<br /><font color="red"><%= ocouponitemlist.FitemList(i).GetFreeBeasongCouponMargin %></font>%
			<% end if %>
		<% if (ocouponitemlist.FitemList(i).Forgprice>ocouponitemlist.FitemList(i).FSellcash) then %>
			<br /><font color="gray">(<%= ocouponitemlist.FitemList(i).GetCouponMarginOrgPrice %></font>%)
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
	<td colspan="13" align="center">
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
set oitemcouponmaster = Nothing
%>
<form name="frmbuf" method="post" action="/admin/shopmaster/itemcoupon_process.asp">
<input type="hidden" name="mode" value="addcouponitemarr">
<input type="hidden" name="itemcouponidx" value="<%= itemcouponidx %>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="couponbuypricearr" value="">
<input type="hidden" name="couponsellcasharr" value="">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="sailyn" value="<%= sailyn %>">
<input type="hidden" name="mwdiv" value="<%= mwdiv %>">
<input type="hidden" name="defaultmargin" value="">

</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
