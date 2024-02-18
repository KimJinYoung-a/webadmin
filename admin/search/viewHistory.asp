<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<!-- #include virtual="/lib/classes/search/searchCls.asp" -->
<%

dim i, j, k
dim userid, page

dim oViewHistory
set oViewHistory = new ViewHistoryCls

userid = oViewHistory.GetRandomUserID()
page = 1
oViewHistory.FPageSize        = 100
oViewHistory.FCurrpage        = page
oViewHistory.FScrollCount     = 10
oViewHistory.FRectUserID      = userid

if userid<>"" then
    oViewHistory.getMyTodayViewListNew
end if

dim currColor, prevTime
dim salePer

%>
<script>
function popRegWord() {
    var window_width = 600;
    var window_height = 350;

    var popwin = window.open("popRegWord.asp","popRegWord","width=" + window_width + " height=" + window_height + " scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsSubmit() {
	document.frm.submit();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			userid : <%= Left(userid, Len(userid) - 2) %>**
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="jsSubmit();">
		</td>
	</tr>
   </form>
</table>

<p />

<input type="button" class="button" value="등록하기" onClick="popRegWord();">

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oViewHistory.FResultCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="180">조회일시</td>
		<td width="60">itemid</td>
		<td width=50> 이미지</td>
		<td width="100">브랜드ID</td>
		<td width="300">카테고리</td>
		<td>상품명</td>
		<td width="120">키워드</td>
		<td width="60">판매가</td>
		<td width="60">할인가</td>
		<td width="80">판매시작일</td>
		<td width="50">후기</td>
		<td>비고</td>
    </tr>
<% if oViewHistory.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oViewHistory.FresultCount > 0 then %>
    <%
	currColor = "#FFFFFF"
	prevTime = oViewHistory.FItemList(i).FRegdate
	for i = 0 to oViewHistory.FresultCount - 1
		if DateDiff("s", prevTime, oViewHistory.FItemList(i).FRegdate) > 10*60 then
			currColor = CHKIIF(currColor="#FFFFFF", "#DDDDFF", "#FFFFFF")
		end if
		prevTime = oViewHistory.FItemList(i).FRegdate
	%>
	<tr class="a" height="25" bgcolor="<%= currColor %>">
		<td align="center"><%= oViewHistory.FItemList(i).FRegdate %></td>
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oViewHistory.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
				<%= oViewHistory.FItemList(i).Fitemid %>
			</a>
		</td>
		<td align="center"><img src="<%= oViewHistory.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oViewHistory.FItemList(i).Fmakerid %></td>
		<td align="left"><%= oViewHistory.FItemList(i).FCateName %></td>
		<td><%= oViewHistory.FItemList(i).FItemName %></td>
		<td align="center"><%= oViewHistory.FItemList(i).Fkeywords %></td>
		<td align="right"><%= FormatNumber(oViewHistory.FItemList(i).FOrgPrice, 0) %></td>
		<td align="right">
			<%= CHKIIF(oViewHistory.FItemList(i).FSaleYn="Y", FormatNumber(oViewHistory.FItemList(i).FSailPrice, 0), "") %>
			<%
			if oViewHistory.FItemList(i).FSaleYn="Y" then
				salePer = (100 - 100*oViewHistory.FItemList(i).FSailPrice/oViewHistory.FItemList(i).FOrgPrice)
			%>
			<br />(<% if salePer >30 then %><font color=red><%= FormatNumber(salePer, 0) %>%</font><% else %><%= FormatNumber(salePer, 0) %>%<% end if %>)
			<% end if %>
		</td>
		<td align="center">
			<%= Left(oViewHistory.FItemList(i).Fsdate, 10) %>
			<% if DateDiff("m", oViewHistory.FItemList(i).Fsdate, Now()) < 3 then %>
			<br /><font color=red>NEW</font>
			<% end if %>
		</td>
		<td align="center"><%= FormatNumber(oViewHistory.FItemList(i).Fevalcnt, 0) %></td>
		<td></td>
	</tr>
	<% next %>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
