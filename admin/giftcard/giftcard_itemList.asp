<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 기프트카드 상품목록
' History : 이상구 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcard_cls.asp"-->
<%
	dim cardItemid, cardItemName, cardSellYN, page, i
	dim oGiftcard

	cardItemid      = requestCheckvar(request("cardItemid"),255)
	cardItemName    = request("cardItemName")
	cardSellYN      = requestCheckvar(request("cardSellYN"),10)
	page 			= requestCheckvar(request("page"),10)

	if (page="") then page=1
	
	if cardItemid<>"" then
		dim iA ,arrTemp,arrItemid
	
		arrTemp = Split(cardItemid,",")
	
		iA = 0
		do while iA <= ubound(arrTemp)
			if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
				arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
			end if
			iA = iA + 1
		loop
	
		if len(arrItemid)>0 then
			cardItemid = left(arrItemid,len(arrItemid)-1)
		else
			if Not(isNumeric(cardItemid)) then
				cardItemid = ""
			end if
		end if
	end if

	set oGiftcard = new cGiftCard
	oGiftcard.FPageSize			= 30
	oGiftcard.FCurrPage			= page
	oGiftcard.FRectCardItemid	= cardItemid
	oGiftcard.FRectSellYn		= cardSellYN

	oGiftcard.fGiftcard_Itemlist
%>
<script type='text/javascript'>
<!--
	//페이지 이동
	function goPage(pg) {
		document.frm.page=pg;
		document.frm.submit();
	}

	// 상품 수정
	function editItemInfo(cardid) {
		if(!cardid) cardid="";
		var pop = window.open("popEditGiftCardItem.asp?cardid="+cardid,"popGiftItem","width=1200,height=700,scrollbars=yes");
		pop.focus();
	}

	// 옵션 수정
	function editGiftOpt(cardid,cardOption) {
		if(!cardOption) cardOption="";
		var pop = window.open("popEditGiftCardOption.asp?cardid="+cardid+"&cardOption="+cardOption,"popGiftOption","width=1200,height=500,scrollbars=yes");
		pop.focus();
	}

	// 디자인 목록
	function popDesignList(cardid) {
		var pop = window.open("popGiftCardDesignList.asp?cardid="+cardid,"popGiftDesign","width=1200,height=700,scrollbars=yes");
		pop.focus();
	}
//-->
</script>

<!-- 검색 시작 -->
<form name="frm" method=get style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상품코드 :
			<input type="text" class="text" name="cardItemid" value="<%= cardItemid %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="cardItemName" value="<%= cardItemName %>" size="32" maxlength="32">
			<br>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "cardSellYN", cardSellYN %>
		</td>
	</tr>
</table>
</form>

<!-- 액션 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding:5px 0 5px 0;"><input type="button" class="button" value="+신규등록" onclick="editItemInfo()"></td>
</tr>
</table>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="6">
			검색결과 : <b><%= oGiftcard.FTotalCount%></b>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oGiftcard.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="60">No.</td>
		<td width=50> 이미지</td>
		<td> 상품명</td>
		<td width="300">옵션</td>
		<td width="30">판매<br>여부</td>
		<td width="50">디자인</td>
    </tr>
<% if oGiftcard.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oGiftcard.FresultCount > 0 then %>
    <% for i=0 to oGiftcard.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center">				
			<a href="<%=WWWUrl%>/shopping/giftcard/giftcard.asp?cardid=<%= oGiftcard.FItemList(i).FcardItemId %>" target="_blank" title="미리보기">				
			<%= oGiftcard.FItemList(i).FcardItemId %></a>
		</td>
		<td align="center"><a href="javascript:editItemInfo('<%= oGiftcard.FItemList(i).FcardItemId %>')" title="기프트카드 상품정보 수정"><img src="<%= oGiftcard.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></a></td>
		<td align="left">
			<a href="javascript:editItemInfo('<% =oGiftcard.FItemList(i).FcardItemId %>')" title="기프트카드 상품정보 수정">
			<%= ReplaceBracket(oGiftcard.FItemList(i).FcardItemName) %></a>
		</td>
		<td align="center"><%=oGiftcard.FItemList(i).fGiftcard_optlist%></td>
		<td align="center"><%= fnColor(oGiftcard.FItemList(i).FcardSellYN,"yn") %></td>
		<td align="center"><a href="javascript:popDesignList('<% =oGiftcard.FItemList(i).FcardItemId %>')" title="기프트카드 디자인정보 수정"><%= oGiftcard.FItemList(i).FdesignCnt %></a></td>
	</tr>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="6" align="center">
			<% if oGiftcard.HasPreScroll then %>
			<a href="javascript:goPage('<%= oGiftcard.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oGiftcard.StartScrollPage to oGiftcard.FScrollCount + oGiftcard.StartScrollPage - 1 %>
    			<% if i>oGiftcard.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oGiftcard.HasNextScroll then %>
    			<a href="javascript:goPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
	
</table>
<% end if %>

<% set oGiftcard = Nothing %>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->