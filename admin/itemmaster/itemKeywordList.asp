<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%

dim page
dim sellyn, usingyn, makerid, itemid, itemname, keyword, searchKey
dim itemidMn, itemidMx

page 		= requestCheckvar(request("page"),10)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
makerid     = requestCheckvar(request("makerid"),32)
itemid     	= requestCheckvar(request("itemid"),32)
itemname    = requestCheckvar(request("itemname"),32)
keyword     = requestCheckvar(request("keyword"),32)
itemidMn    = requestCheckvar(request("itemidMn"),32)
itemidMx    = requestCheckvar(request("itemidMx"),32)
searchKey   = requestCheckvar(request("searchKey"),32)

if (page="") then page=1


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 100
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemidMin    = itemidMn
oitem.FRectItemidMax    = itemidMx

oitem.FRectSearchKey    = searchKey

if (makerid <> "") or (itemidMn <> "" and itemidMx <> "") then
	oitem.FRectItemName     = itemname
	oitem.FRectKeyword 		= keyword
else
	if (itemname <> "" or keyword <> "") then
		response.write "<script>alert('먼저 브랜드 또는 상품코드 대역을 지정하세요.');</script>"
	end if
end if
oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn

oitem.GetItemKeywordList


dim i

%>
<script>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function jsSubmit() {
	var frm = document.frm;

	if (frm.itemidMn.value != "") {
		if (frm.itemidMn.value*0 != 0) {
			alert("숫자만 가능합니다.");
			frm.itemidMn.focus();
			return;
		}
	}

	if (frm.itemidMx.value != "") {
		if (frm.itemidMx.value*0 != 0) {
			alert("숫자만 가능합니다.");
			frm.itemidMx.focus();
			return;
		}
	}

	if ((frm.itemname.value != "") || (frm.keyword.value != "")) {
		if ((frm.makerid.value == "") && ((frm.itemidMn.value == "") || (frm.itemidMx.value == ""))) {
			alert("상품명 또는 키워드 검색을 할 경우\n\n브랜드 또는 상품코드 대역을 지정해야 합니다.");
			return;
		}

		if ((frm.itemidMn.value != "") && (frm.itemidMx.value != "")) {
			if ((frm.itemidMx.value*1 - frm.itemidMn.value*1) > 10000) {
				alert("상품코드 대역은 최대 1만번대 까지 가능합니다.");
				return;
			}
		}
	}

	document.frm.submit();
}

function popXL() {
	<% if (oitem.FTotalCount > 10000) then %>
	alert("최대 1만개까지 가능합니다.");
	return;
	<% end if %>

	var makerid = "<%= makerid %>";
	var sellyn = "<%= sellyn %>";
	var usingyn = "<%= usingyn %>";
	var itemid = "<%= itemid %>";
	var itemname = "<%= itemname %>";
	var keyword = "<%= keyword %>";
	var itemidMn = "<%= itemidMn %>";
	var itemidMx = "<%= itemidMx %>";
	var searchKey = "<%= searchKey %>";

	var popwin = window.open("itemKeyword_xl_download.asp?makerid=" + makerid + "&sellyn=" + sellyn + "&usingyn=" + usingyn + "&itemid=" + itemid + "&itemname=" + itemname + "&keyword=" + keyword + "&itemidMn=" + itemidMn + "&itemidMx=" + itemidMx + "&searchKey=" + searchKey,"popXL","width=100,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popItemKeywordModify(itemid) {
	var popwin = window.open("pop_itemKeywordModify.asp?itemid=" + itemid,"popItemKeywordModify","width=1000,height=200 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popItemKeywordModifyMulti() {
	var popwin = window.open("pop_itemKeywordModifyMulti.asp","popItemKeywordModifyMulti","width=1000,height=500 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드: <%	drawSelectBoxDesignerWithName "makerid", makerid %>
			판매여부 : <% drawSelectBoxSellYN "sellyn", sellyn %>
			사용여부 : <% drawSelectBoxUsingYN "usingyn", usingyn %>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="jsSubmit();">
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			상품코드: <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="16">
			상품명: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
			키워드: <input type="text" class="text" name="keyword" value="<%= keyword %>" size="32" maxlength="32">
			&nbsp;
			상품코드 대역:
			<input type="text" class="text" name="itemidMn" value="<%= itemidMn %>" size="10" maxlength="32">
			-
			<input type="text" class="text" name="itemidMx" value="<%= itemidMx %>" size="10" maxlength="32">
			(예 : 1000000 - 1050000)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td align="left">
			상품명+키워드: <input type="text" class="text" name="searchKey" value="<%= searchKey %>" size="32" maxlength="32"> (최대 5000개까지 검색됩니다.)
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr>
    <td>

	</td>
	<td align="right">
		<input type="button" class="button" value="일괄적용" onclick="popItemKeywordModifyMulti();">
		&nbsp;
		<input type="button" class="button" value="엑셀받기" onclick="popXL();">
    </td>
</tr>
</table>

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oitem.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/<%=  oitem.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">IDX</td>
		<td width="60">itemID</td>
		<td width=50> 이미지</td>
		<td width="100">브랜드ID</td>
		<td>상품명</td>
		<td>키워드</td>
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td>비고</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td align="center"><%= oitem.FTotalCount - (i + (page - 1)*100) %></td>
		<td align="center">
			<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank" title="미리보기">
			<%= oitem.FItemList(i).Fitemid %></a>
			</td>
		<td align="center"><img src="<%= oitem.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
		<td align="left"><%= oitem.FItemList(i).Fmakerid %></td>
		<td align="left"><%= oitem.FItemList(i).Fitemname %></td>
		<td align="left"><%= oitem.FItemList(i).Fkeywords %></td>
		<td align="center"><%= oitem.FItemList(i).Fsellyn %></td>
		<td align="center"><%= oitem.FItemList(i).Fisusing %></td>
		<td align="center">
			<input type="button" class="button" value="수정" onClick="popItemKeywordModify(<%= oitem.FItemList(i).Fitemid %>)">
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

</table>
<% end if %>

<%
SET oitem = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
