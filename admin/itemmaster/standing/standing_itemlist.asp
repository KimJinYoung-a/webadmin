<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정기구독관리
' History : 2016.06.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim cdl, cdm, cds, dispCate, page, reloading
dim deliverytype, i
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
	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)
	dispCate = requestCheckvar(request("disp"),16)
	reloading = requestCheckvar(request("reloading"),2)
	page = requestCheckvar(request("page"),10)

if (page="") then page=1
if reloading="" and itemdiv="" then itemdiv="75"

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
	oitem.FRectDispCate		= dispCate
	oitem.FRectItemDiv      = itemdiv
	oitem.FRectDeliverytype = deliverytype
	oitem.GetItemoptionList

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function NextPage(ipage){
	document.frm.page.value= ipage;
	if ((document.frm.itemname.value.length>0)&&(ipage*1==1)){
	    alert('상품명 검색시 결과는 최대 1000개로 제한됩니다.');  // 2차서버 fulltext 검색후 조인방식으로 변경.
	}
	
	document.frm.submit();
}

//정기구독상품 정보 수정
function editstandingItemInfo(itemid,itemoption) {
	var param = "itemid=" + itemid + "&itemoption=" + itemoption + "&makerid=<%= makerid %>&page=<%= page %>&menupos=<%= menupos %>";
	window.open('<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standingIteminfo.asp?' + param ,'editstandingItemInfo','width=1280,height=768,scrollbars=yes,resizable=yes');
}

// 정기구독 주문리스트
function standingorderchulgo() {
	var param = "sendstatus=05&menupos=<%= menupos %>";
	window.open('<%= getSCMSSLURL %>/admin/itemmaster/standing/pop_standinguser.asp?' + param ,'standingorderchulgo','width=1280,height=960,scrollbars=yes,resizable=yes');
}

</script>

<form name="frm" method="get" style="margin:0;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="reloading" value="ON">
<input type="hidden" name="page" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td style="white-space:nowrap;">브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %> </td>
				<td style="white-space:nowrap;padding-left:5px;">상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"> </td>
				<td style="white-space:nowrap;padding-left:5px;">상품코드:</td> 
				<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
		</tr> 
		<tr>
			<td  style="white-space:nowrap;">관리<!-- #include virtual="/common/module/categoryselectbox.asp"--> </td> 
			<td  style="white-space:nowrap;"  colspan="2">&nbsp;&nbsp;전시카테고리: <!-- #include virtual="/common/module/dispCateSelectBox.asp"--> </td>
			<td ></td>
		</tr>
		</table>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="NextPage(1);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<span style="white-space:nowrap;">판매:<% drawSelectBoxSellYN "sellyn", sellyn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">사용:<% drawSelectBoxUsingYN "usingyn", usingyn %></span>
     	&nbsp;     	
     	<span style="white-space:nowrap;">단종:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">한정:<% drawSelectBoxLimitYN "limityn", limityn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">과세: <% drawSelectBoxVatYN "vatyn", vatyn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">할인 <% drawSelectBoxSailYN "sailyn", sailyn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">해외배송 <% drawSelectBoxIsOverSeaYN "overSeaYn", overSeaYn %></span>
     	&nbsp;
     	<span style="white-space:nowrap;">배송구분 <% drawBeadalDiv "deliverytype", deliverytype %></span>
        &nbsp;
     	<span style="white-space:nowrap;">상품구분 <% drawSelectBoxItemDiv "itemdiv", itemdiv %></span>
	</td>
</tr>
</table>
</form>

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" onclick="standingorderchulgo();" value="출고리스트" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">No.</td>
	<td width="60">옵션코드</td>
	<td width=50>이미지</td>
	<td width="100">브랜드ID</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="60">판매가</td>
	<td width="30">계약<br>구분</td>
	<td width="30">판매<br>여부</td>
	<td width="30">사용<br>여부</td>
	<td width="30">한정<br>여부</td>
	<td width="50">옵션<br>사용여부</td>
	<td width="50">옵션<br>판매여부</td>
	<td width="90">비고</td>
</tr>

<% if oitem.FresultCount>0 then %>
	<% for i=0 to oitem.FresultCount-1 %>
	<tr bgcolor="<%=chkIIF(oitem.FItemList(i).foptisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(oitem.FItemList(i).foptisusing="Y","#FFFFFF","#DDDDDD")%>';>
		<td align="center"><%= oitem.FItemList(i).Fitemid %></td>
		<td align="center"><%= oitem.FItemList(i).fitemoption %></td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left">
			<% =oitem.FItemList(i).Fitemname %>
	
			<% if oitem.FItemList(i).FitemDiv="08" then %>
				<font color="#F89020">[Ticket]</font>
			<% end if %>
			<% if oitem.FItemList(i).FitemDiv="18" then %>
				<font color="#F89020">[travel]</font>
			<% end if %>
		</td>
		<td align="left">
			<% =oitem.FItemList(i).foptionname %>
		</td>
		<td align="right">
		<%
			Response.Write FormatNumber(oitem.FItemList(i).Forgprice,0)
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
		<td align="center">
			<%= fnColor(oitem.FItemList(i).Fmwdiv,"mw") %>
			<br>
			<%
				If oitem.FItemList(i).Fdeliverytype = "1" Then
					response.write "텐배"
				ElseIf oitem.FItemList(i).Fdeliverytype = "2" Then
					response.write "무료"
				ElseIf oitem.FItemList(i).Fdeliverytype = "4" Then
					response.write "텐무"
				ElseIf oitem.FItemList(i).Fdeliverytype = "9" Then
					response.write "조건"
				ElseIf oitem.FItemList(i).Fdeliverytype = "7" Then
					response.write "착불"
				End If
			%>
		</td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Fisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).Flimityn,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).foptisusing,"yn") %></td>
		<td align="center"><%= fnColor(oitem.FItemList(i).foptsellyn,"yn") %></td>
		<td align="center">
			<% if oitem.FItemList(i).FitemDiv="75" then %>
				<input type="button" onclick="editstandingItemInfo('<% =oitem.FItemList(i).Fitemid %>','<% =oitem.FItemList(i).fitemoption %>');" value="정기구독등록" class="button">
			<% end if %>
		</td>
	</tr>
	<% next %>
	
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
<% else %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>

<%
SET oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->