<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
Response.CharSet = "euc-kr"

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown, dispCate, sortdiv, searchgubun, searchtxt
dim page
dim infodivYn, infodiv, deliverytype

searchgubun = requestCheckvar(request("searchgubun"),10)
searchtxt	 = requestCheckvar(request("searchtxt"),100)
itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
limityn     = requestCheckvar(request("limityn"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
sailyn      = requestCheckvar(request("sailyn"),10)
overSeaYn   = requestCheckvar(request("overSeaYn"),10)
itemdiv     = requestCheckvar(request("itemdiv"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
sortdiv		 = NullFillWith(requestCheckvar(request("sortdiv"),5),"new")

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

showminusmagin = request("showminusmagin")
marginup = request("marginup")
margindown = request("margindown")

infodiv  = request("infodiv")
infodivYn  = requestCheckvar(request("infodivYn"),10)

If infodiv <> "" Then
	infodivYn = "Y"	
End If

page = requestCheckvar(request("page"),10)

if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 10
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid

If searchgubun = "itemname" Then
	oitem.FRectItemName     = searchtxt
ElseIf searchgubun = "itemid" Then
	oitem.FRectItemid       = searchtxt
End IF

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectVatYn        = vatyn
oitem.FRectSailYn       = sailyn
oitem.FRectIsOversea	= overSeaYn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectItemDiv      = itemdiv

oitem.FRectMinusMigin = showminusmagin
oitem.FRectMarginUP = marginup
oitem.FRectMarginDown = margindown
oitem.FRectInfodivYn    = infodivYn
oitem.FRectInfodiv    = infodiv 
oitem.FRectDeliverytype = deliverytype
oitem.FRectSortDiv	= sortdiv
oitem.GetItemList

dim i
%>
<form id="itemfrm" name="itemfrm" method="get" style="margin:0px;">
<input type="hidden" id="page" name="page">
<div class="searchWrap" style="border-top:none;">
	<div class="search">
		<ul>
			<li>
				<label class="formTit">카테고리 :</label>
				<!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			</li>
		</ul>
	</div>
	<dfn class="line"></dfn>
	<div class="search">
		<ul>
			<li>
				<label class="formTit" for="schWord">검색어 :</label>
				<select class="formSlt" id="searchgubun" name="searchgubun" title="옵션 선택">
					<option value="itemname" <%=CHKIIF(searchgubun="itemname","selected","")%>>상품명</option>
					<option value="itemid" <%=CHKIIF(searchgubun="itemid","selected","")%>>상품ID</option>
				</select>
				<input type="text" class="formTxt" id="searchtxt" name="searchtxt" value="<%=searchtxt%>" style="width:400px" placeholder="상품명 또는 상품ID를 입력하여 검색하세요." maxlength="100" onKeyPress="if (event.keyCode == 13){ NextPage(1,'item'); return false;}" />
			</li>
		</ul>
	</div>
	<input type="button" id="btnsearh1" class="schBtn" value="검색" onClick="NextPage(1,'item');" />
	<input type="button" id="btnsearh2" style="display:none;" class="schBtn" value="검색" onClick="alert('검색중입니다. 잠시만 기다려주세요.');" />
</div>
<br />
<div class="tbListWrap">
	<!--
	<div class="ftLt lPad10">
		<select class="formSlt" id="sortdiv" name="sortdiv" title="옵션 선택" onChange="jsSortDiv('item');">
			<option value="new" <%=CHKIIF(sortdiv="new","selected","")%>>신상품순</option>
			<option value="best" <%=CHKIIF(sortdiv="best","selected","")%>>인기순</option>
		</select>
	</div>
	//-->
	<div class="ftLt lPad10"></div>
	<div class="ftRt pad10">
		<span>검색결과 : <strong><%= oitem.FTotalCount%></strong></span> <span class="lMar10">페이지 : <strong><%= page %> / <%=  oitem.FTotalpage %></strong></span>
	</div>
</div>
</form>

<div class="tbListWrap tMar15">

	<ul class="thDataList">
		<li>
			<p class="cell05"></p>
			<p class="cell10">상품 ID</p>
			<p class="cell10">이미지</p>
			<p>상품명</p>
			<p class="cell10">가격</p>
			<p class="cell10">업체 ID</p>
			<p class="cell10">판매여부</p>
		</li>
	</ul>
	<ul class="tbDataList" id="contentslist">
	<% if oitem.FresultCount > 0 then %>
	    <% for i=0 to oitem.FresultCount-1 %>
		<li id="tr<%= oitem.FItemList(i).Fitemid %>" style="cursor:pointer;">
			<p class="cell05"><input type="checkbox" name="contentsidx<%= oitem.FItemList(i).Fitemid %>" id="contentsidx<%= oitem.FItemList(i).Fitemid %>" value="<%= oitem.FItemList(i).Fitemid %>" onClick="jsThisCheck('<%=oitem.FItemList(i).Fitemid%>','item');" /></p>
			<p class="cell10" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');"><%= oitem.FItemList(i).Fitemid %></p>
			<p class="cell10" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0" /></p>
			<p class="lt" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');"><% =oitem.FItemList(i).Fitemname %></p>
			<p class="cell10" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');">
				<%
					if oitem.FItemList(i).Fsailyn="Y" then
						response.write FormatNumber(oitem.FItemList(i).Fsailprice,0)
					else
						response.write FormatNumber(oitem.FItemList(i).Forgprice,0)
					end if
				%>
			</p>
			<p class="cell10" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');"><%= oitem.FItemList(i).Fmakerid %></p>
			<p class="cell10" onClick="jsThisClick('<%= oitem.FItemList(i).Fitemid %>','item');"><%=oitem.FItemList(i).Fsellyn%></p>
		</li>
		<% next %>
	<% end if %>
	</ul>
	<div class="ct tPad20 cBk1 bPad10">
		<% if oitem.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>','item')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<span class="cRd1">[<%= i %>]</span>
			<% else %>
			<a href="javascript:NextPage('<%= i %>','item')">[<%= i %>]</a>
			<% end if %>
		<% next %>
		<% if oitem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>','item')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</div>
</div>
<% SET oitem = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->