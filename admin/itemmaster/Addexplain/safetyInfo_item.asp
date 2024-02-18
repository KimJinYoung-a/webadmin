<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품 안정인증정보 관리 - 상품목록
' Hieditor : 2015.05.28 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown
dim page
dim infodivYn, saftyYn, saftyInfoYn

itemid      = requestCheckvar(request("itemid"),255)
itemname    = request("itemname")
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

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

showminusmagin = request("showminusmagin")
marginup = request("marginup")
margindown = request("margindown")
infodivYn  = requestCheckvar(request("infodivYn"),10)
saftyYn  = requestCheckvar(request("saftyYn"),1)
saftyInfoYn  = requestCheckvar(request("saftyInfoYn"),1)
If saftyInfoYn="" Then saftyInfoYn = "N"
''If infodivYn = "K" Then sellyn = "Y"

If marginup <> "" AND IsNumeric(marginup) = False Then
	rw "<script>alert('마진값(이상)이 잘못되었습니다. - "&marginup&"');history.back();</script>"
	dbget.close()
	Response.End
End If

If margindown <> "" AND IsNumeric(margindown) = False Then
	rw "<script>alert('마진값(이하)이 잘못되었습니다. - "&margindown&"');history.back();</script>"
	dbget.close()
	Response.End
End If

page = requestCheckvar(request("page"),10)

if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

oitem.FPageSize         = 30
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

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
oitem.FRectItemDiv      = itemdiv

oitem.FRectMinusMigin = showminusmagin
oitem.FRectMarginUP = marginup
oitem.FRectMarginDown = margindown
oitem.FRectInfodivYn    = infodivYn
oitem.FRectsaftyYn		= saftyYn
oitem.FRectsaftyInfoYn  = saftyInfoYn
oitem.FRectShowInfoDiv  = "on"
oitem.FRectSortDiv="best"               ''베스트순.

oitem.getSafetyInfoItemList

dim i
%>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

// ============================================================================
// 옵션수정 -교체
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//판매수정
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

// ============================================================================
// 이미지수정
function editItemImage(itemid, makerid) {
	var param = "itemid=" + itemid;
	
	//if(makerid =="ithinkso"){
		//popwin = window.open('/common/pop_itemimage_ithinkso.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}else{
		popwin = window.open('/common/pop_itemimage.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	//}
	popwin.focus();
}

// 상품설명 이미지 등록/수정
function popItemContImage(itemid)
{
	var popwin = window.open("/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=" + itemid + "&menupos=423","popitemContImage","width=600 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 재고현황 팝업
function PopItemStock(itemid){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemid=" + itemid,"popitemstocklist","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 기본정보 수정
function editItemBasicInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemBasicInfo.asp?' + param ,'editItemBasic','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// 판매가 및 공급가 설정
function editItemPriceInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ItemPriceInfo.asp?' + param ,'editItemPrice','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//티켓 상품 정보 수정
function editTicketItemInfo(itemid) {
	var param = "itemid=" + itemid + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	popwin = window.open('/admin/itemmaster/pop_ticketIteminfo.asp?' + param ,'pop_ticketIteminfo','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//과세,면세 수정 팝업
function vatedit(itemid,vat){
	var param = "itemid=" + itemid + "&vat="+vat+"";
	popwin = window.open('/admin/itemmaster/pop_vatEdit.asp?' + param ,'pop_vatEdit','width=300,height=150');
	popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
			<br>
		</td>
		
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;     	
	     	단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	     	&nbsp;
	     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
	     	과세: <% drawSelectBoxVatYN "vatyn", vatyn %>
	     	&nbsp;
	     	할인 <% drawSelectBoxSailYN "sailyn", sailyn %>
	     	
	     	&nbsp;
	     	해외배송 <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %>
            &nbsp;
	     	상품구분 <% drawSelectBoxItemDiv "itemdiv", itemdiv %>
	     	&nbsp;
	     	품목정보입력여부
			<select class="select" name="infodivYn">
			<option value="">전체</option>
			<option value="N" <%= CHKIIF(infodivYn="N","selected","") %> >입력이전</option>
			<option value="Y" <%= CHKIIF(infodivYn="Y","selected","") %> >입력완료</option>
			<option value="K" <%= CHKIIF(infodivYn="K","selected","") %> >항목누락</option>
			</select>
			&nbsp;
			<font color="red">안전인증대상여부</font>
			<select class="select" name="saftyYn">
			<option value="">전체</option>
			<option value="N" <%= CHKIIF(saftyYn="N","selected","") %> >대상아님</option>
			<option value="Y" <%= CHKIIF(saftyYn="Y","selected","") %> >인증대상</option>
			</select>
			&nbsp;
			<font color="red">KC마크입력여부</font>
			<select class="select" name="saftyInfoYn">
			<option value="A" <%= CHKIIF(saftyInfoYn="A","selected","") %> >전체</option>
			<option value="N" <%= CHKIIF(saftyInfoYn="N","selected","") %> >입력이전</option>
			<option value="Y" <%= CHKIIF(saftyInfoYn="Y","selected","") %> >입력완료</option>
			</select>
    </tr>
    </form>
</table>

<p>
<% If cdl = "110" and cdm = "010" and cds = "968" Then %>
<input type="button" value="포토북 템플릿코드 등록" class="button" onClick="window.open('pop_photobook.asp','popPhotobook','width=600,height=650,scrollbars=yes');"><p>
<% End If %>

<!-- 리스트 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			검색결과 : <b><%= oitem.FTotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">No.</td>
		<td width=50> 이미지</td>
		<td width="100">브랜드ID</td>
		<td> 상품명</td>
		<td width="60">판매가</td>
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="60" bgcolor="#FFF0F0">안전인증<br>여부</td>
		<td width="130" bgcolor="#FFF0F0">안전인증<br>구분</td>
		<td width="170" bgcolor="#FFF0F0">안전인증<br>번호</td>
		<td width="150">품목</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="16" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">				
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">				
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>','<%= oitem.FItemList(i).Fmakerid %>')" title="이미지 수정"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></a></td>
		<td align="left"><a href="javascript:PopBrandInfoEdit('<%= oitem.FItemList(i).Fmakerid %>')" title="브랜드 정보 수정"><%= oitem.FItemList(i).Fmakerid %></a></td>
		<td align="left">
			<a href="javascript:editItemBasicInfo('<% =oitem.FItemList(i).Fitemid %>')" title="상품 기본정보 수정"><% =oitem.FItemList(i).Fitemname %></a>
			<% if oitem.FItemList(i).FitemDiv="08" then %>
            <a href="javascript:editTicketItemInfo('<% =oitem.FItemList(i).Fitemid %>')" title="Ticket 정보 수정"><font color="#F89020">[Ticket]</font></a>	
			<% end if %>
		</td>
		<td align="right">
		<%
			Response.Write "<a href=""javascript:editItemPriceInfo('" & oitem.FItemList(i).Fitemid & "')"" title='판매가 및 공급가 설정'>" & FormatNumber(oitem.FItemList(i).Forgprice,0) & "</a>"
			'할인가
			if oitem.FItemList(i).Fsailyn="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
			end if
			'쿠폰가
			if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
					Case "2"
						Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				end Select
			end if
		%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>

		<td align="center"><%= fnColor(oitem.FItemList(i).FsafetyYn,"yn") %></td>
	    <td align="center"><%= getSaftyDivName(oitem.FItemList(i).FsafetyYn,oitem.FItemList(i).FsafetyDiv) %></td>
	    <td align="center"><%= oitem.FItemList(i).FsafetyNum %></td>
	    <td align="center"><%= getAddExpInfoDivName(oitem.FItemList(i).FinfoDiv) %></td>
	</tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	
</table>
<% end if %>

<%
SET oitem = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->