<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 다이 상품 등록 대기 상품 
' Hieditor : 2010.10.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/diyshopitem/waitDIYitemCls.asp"-->

<%
Dim owaititem,ix,page ,sorttype, sortkey, sortkeyMid, currstate
	page = RequestCheckvar(request("page"),10)
	currstate = RequestCheckvar(request("currstate"),10)
	sorttype  = RequestCheckvar(request("sorttype"),10)
	sortkey = RequestCheckvar(request("sortkey"),32)
	sortkeyMid = RequestCheckvar(request("sortkeyMid"),10)
	
	if (page="") then page=1
	
	if sorttype="" then sorttype="C"
	if currstate="" then currstate="W"

set owaititem = new CWaitItemlist
	owaititem.FPageSize = 30
	owaititem.FCurrPage = page
	owaititem.FRectsortkey = sortkey
	owaititem.FRectsortkeyMid = sortkeyMid
	owaititem.FRectCurrState = currstate
	
	if sorttype="C" then
		owaititem.getWaitProductListByCategory
	elseif sorttype="B" then
		owaititem.getWaitProductListByBrand
	end if
%>

<script language='javascript'>

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ViewItemDetail(itemno){
	window.open('/academy/itemmaster/viewDIYitem/viewDIYitem.asp?itemid='+itemno ,'window1','width=1024,height=960,scrollbars=yes,status=no');
}

function insertdb(itemid,itemname){
 //if (confirm(itemname + "를 등록하시겠습니까?") == true){
    //location.href("item_insertdb.asp?itemid="+itemid);
 //}
}

function WaitState(itemid){
	var ret = confirm('등록대기로 변경하시겠습니까?');

	if (ret){
		document.location = 'doitemregboru.asp?mode=waitstate&idx=' + itemid;
	}
}

function popItemModify(itemid,designer){
	var popwin = window.open('wait_diyitem_modify.asp?itemid=' + itemid + '&designer=' + designer +'&fingerson=on','waititemmodify','width=860,height=700,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="sorttype" value="<%= sorttype %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input type="radio" name="currstate" value="W" <% if currstate="W" then response.write "checked" %>>등록대기상품만
		<input type="radio" name="currstate" value="WR" <% if currstate="WR" then response.write "checked" %>>등록대기+등록보류
		<input type="radio" name="currstate" value="A" <% if currstate="A" then response.write "checked" %>>전체
		&nbsp;
		<% if sorttype="C" then %>
			카테고리 :
			<% DrawSelectBoxCategoryLarge "sortkey" , sortkey %>&nbsp;
			<% DrawSelectBoxCategoryMid "sortkeyMid" , sortkey, sortkeyMid %>
		<% else %>
			브랜드 :
			<% drawSelectBoxLecturer "sortkey" , sortkey %>
		<% end if %>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">			
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<br>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if owaititem.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= owaititem.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= owaititem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">No.</td>
	<td align="center">상품명</td>
	<td align="center">미리보기</td>
	<td align="center">판매가</td>
	<td align="center">공급가</td>
	<td align="center">마진</td>
	<td align="center">디자이너</td>
	<td align="center">등록일</td>
	<td align="center">상태</td>
</tr>
<% for ix=0 to owaititem.FresultCount-1 %>

<tr align="center" bgcolor="#FFFFFF" >
	<td align="center"><%= owaititem.FItemList(ix).Fitemid %></td>
	<td align="left"><a href="javascript:popItemModify('<% =owaititem.FItemList(ix).Fitemid %>','<%= owaititem.FItemList(ix).Fmakerid %>')"><%= owaititem.FItemList(ix).Fitemname %></a></td>
	<td align="center">
		<% if owaititem.FItemList(ix).FCurrState="7" then %>
			<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<% =owaititem.FItemList(ix).Flinkitemid %>" target="_blank"><font color="blue">(보기)</font></a>
		<% else %>
			<a href="javascript:ViewItemDetail('<% =owaititem.FItemList(ix).Fitemid %>')"><font color="blue">(미리보기)</font></a>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(owaititem.FItemList(ix).Fsellcash,0) %></td>
	<td align="right"><%= FormatNumber(owaititem.FItemList(ix).Fsuplycash,0) %></td>
	<td align="center">
	<% if owaititem.FItemList(ix).Fsellcash<>0 then %>
	<%= 100 - CLng(owaititem.FItemList(ix).Fsuplycash/owaititem.FItemList(ix).Fsellcash*100*100)/100 %> %
	<% end if %>
	</td>
	<td align="center"><%= owaititem.FItemList(ix).Fmakerid %></td>
	<td align="center"><%= owaititem.FItemList(ix).Fregdate %></td>
	<td align="center"><font color="<%= owaititem.FItemList(ix).GetCurrStateColor %>"><%= owaititem.FItemList(ix).GetCurrStateName %></font>
	<% if (owaititem.FItemList(ix).FCurrState="2") or (owaititem.FItemList(ix).FCurrState="0") then %>
	<a href="javascript:WaitState('<%= owaititem.FItemList(ix).Fitemid %>')"><br><font color="#000000">[등록대기변경]</font></a>
	<% end if %>
	</td>
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if owaititem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= owaititem.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + owaititem.StarScrollPage to owaititem.StarScrollPage + owaititem.FScrollCount - 1 %>
			<% if (ix > owaititem.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(owaititem.FCurrPage) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>
	
		<% if owaititem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
<%
	set owaititem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->