<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemImageTextCls.asp"-->
<%

dim page, i, research
dim itemid, makerid

page		= requestCheckvar(Request("page"), 10)
research	= requestCheckvar(Request("research"), 10)
itemid		= requestCheckvar(Request("itemid"), 10)
makerid		= requestCheckvar(Request("makerid"), 32)

if page="" then
	page = 1
end if


'// ============================================================================
dim oitem
set oitem = new CItemImageText

oitem.FPageSize		= 20
oitem.FCurrPage		= page
oitem.FRectMakerId	= makerid
if IsNumeric(itemid) then
	oitem.FRectItemId	= itemid
elseif (itemid <> "") then
	response.write "<script>alert('잘못된 상품코드입니다.')</script>"
end if

oitem.GetItemImageTextList

%>
<script>

function GoPage(ipage){
	document.frm.page.value = ipage;
	document.frm.submit();
}

function jsPopIns() {
	var v = "popItemImageText_Ins.asp";
	var popwin = window.open(v,"jsPopIns","width=150,height=300,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function jsPopModi(itemid) {
	var v = "popItemImageText_Modi.asp?itemid=" + itemid;
	var popwin = window.open(v,"jsPopModi","width=1200,height=800,scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			* 상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="15" maxlength="15" onKeyPress="if (event.keyCode == 13) document.frm.submit();">
			&nbsp;
			* 브랜드 :
			<%	drawSelectBoxDesignerWithName "makerid", makerid %>
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
    </form>
</table>

<p />

<input type="button" class="button" value="요청등록" onClick="jsPopIns()">

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11">
			검색결과 : <b><%= oitem.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> /<%=  oitem.FTotalPage %></b>
		</td>
	</tr>
	</form>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">이미지</td>
		<td>브랜드</td>
		<td width="100">상품코드</td>
		<td>상품명</td>
		<td>상품이미지<br />텍스트</td>
		<td>수정<br />텍스트</td>
		<td width="80">요청일자</td>
		<td width="80">텍스트<br />추출일자</td>
		<td width="120">요청자</td>
		<td width="150">요청일시</td>
		<td>비고</td>
    </tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="9" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
<% if oitem.FResultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><img src="<%= oitem.FItemList(i).FsmallImage %>" border="0" width="50"></td>
		<td bgcolor="#FFFFFF"><%= oitem.FItemList(i).FmakerId %></td>
		<td bgcolor="#FFFFFF"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><%= oitem.FItemList(i).Fitemid %></a></td>
		<td bgcolor="#FFFFFF"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><%= oitem.FItemList(i).FitemName %></a></td>
		<td bgcolor="#FFFFFF"><a href="javascript:jsPopModi(<%= oitem.FItemList(i).Fitemid %>)"><%= Left(oitem.FItemList(i).Fimagetext,200) %><%= CHKIIF(Len(oitem.FItemList(i).Fimagetext)>200, "...", "") %></a></td>
		<td bgcolor="#FFFFFF"><a href="javascript:jsPopModi(<%= oitem.FItemList(i).Fitemid %>)"><%= Left(oitem.FItemList(i).Fmodifiedtext,200) %><%= CHKIIF(Len(oitem.FItemList(i).Fmodifiedtext)>200, "...", "") %></a></td>
		<td bgcolor="#FFFFFF"><%= oitem.FItemList(i).Freq_yyyymmdd %></td>
		<td bgcolor="#FFFFFF"><%= oitem.FItemList(i).Ffin_yyyymmdd %></td>
		<td bgcolor="#FFFFFF"><%= oitem.FItemList(i).Flastuserid %></td>
		<td bgcolor="#FFFFFF"><%= oitem.FItemList(i).Flastupdate %></td>
		<td bgcolor="#FFFFFF"></td>
    </tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11" align="center">
			<% if oitem.HasPreScroll then %>
			<a href="javascript:GoPage('<%= oitem.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
    			<% if i>oitem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:GoPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oitem.HasNextScroll then %>
    			<a href="javascript:GoPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>

</table>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
