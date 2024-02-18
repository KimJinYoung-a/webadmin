<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/PlusSaleItemCls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%
dim makerid, itemidArr, itemname, page
dim cdl, cdm, cds
'dim sellyn,usingyn

page        = RequestCheckVar(request("page"),9)
makerid     = RequestCheckVar(request("makerid"),32)
itemidArr   = RequestCheckVar(request("itemidArr"),1024)
itemname    = RequestCheckVar(request("itemname"),64)
cdl         = RequestCheckVar(request("cdl"),3)
cdm         = RequestCheckVar(request("cdm"),3)
cds         = RequestCheckVar(request("cds"),3)

if (page="") then page=1
itemidArr = Trim(itemidArr)
itemname  = Trim(itemname)
if (Right(itemidArr,1)=",") then itemidArr = Left(itemidArr,Len(itemidArr)-1)


dim oPsItemList
set oPsItemList = new CPlusSaleItem
oPsItemList.FPageSize     = 20
oPsItemList.FCurrPage     = page
oPsItemList.FRectMakerid  = makerid
oPsItemList.FRectCDL      = cdl
oPsItemList.FRectCDM      = cdm
oPsItemList.FRectCDS      = cds
oPsItemList.FRectItemIDArr= itemidArr
oPsItemList.FRectItemName = itemname

oPsItemList.GetPlusSaleMainItemList


dim i
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function PlusSaleItem_Main_New(){
    var popwin = window.open('/admin/shopmaster/PlusSaleItem_Edit.asp','PlusSaleItem_Main_New','');
    popwin.focus();
}

function showLinkedItemList(iitemid){
    var popwin = window.open('PlusSaleItem_Edit.asp?itemid=' + iitemid,'PlusSaleItem_Edit','');
    popwin.focus();
}


</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 :<%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;
			<!-- #include virtual="/common/module/categoryselectbox.asp"-->
			<br>
			상품코드 :
			<input type="text" class="text" name="itemidArr" value="<%= itemidArr %>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.frm.submit();">(쉼표로 복수입력가능)
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!-- <input type="button" class="button" value="신규등록1" onClick="PlusSaleItem_Edit();"> -->
			<input type="button" class="button" value="신규등록" onClick="PlusSaleItem_Main_New();">
		</td>
		<td align="right">
		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->
<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= oPsItemList.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oPsItemList.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">상품코드</td>
    	<td width="50">이미지</td>
      	<td width="100">브랜드ID</td>
      	<td>상품명</td>
      	<td width="60">판매가</td>
		<td width="60">매입가</td>
		<td width="40">마진</td>
		<td width="30">계약<br>구분</td>
      	<td width="100">진행중인<br>추가구성상품수</td>
    </tr>
    <% for i=0 to oPsItemList.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %></td>
    	<td><img src="<%= oPsItemList.FItemList(i).Fsmallimage %>" width="50" height="50" ></td>
      	<td><%= oPsItemList.FItemList(i).FMakerid %></td>
      	<td align="left"><%= oPsItemList.FItemList(i).FitemName %></td>
      	<td align="right">
      	    <%= FormatNumber(oPsItemList.FItemList(i).FOrgPrice,0) %>
          	<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
          		<br><font color=#F08050>(할)<%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %></font>
          	<% end if %>
      	
      	    <% if oPsItemList.FItemList(i).IsCouponItem then %>
      	        <br><font color=#5080F0>(쿠)<%= FormatNumber(oPsItemList.FItemList(i).GetCouponAssignPrice,0) %></font>
      	    <% end if %>
      	</td>
      	<td align="right">
      		<%= FormatNumber(oPsItemList.FItemList(i).Forgsuplycash,0) %> 
      		<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050>(할)<%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %></font>
      	    <% end if %>
      	    
      	</td>
      	<td>
      		<%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %>
      		<% if oPsItemList.FItemList(i).IsCurrentSaleItem then %>
      		<br><font color=#F08050><%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %></font>
      	    <% end if %>
      	</td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FMwdiv,"mw") %></td>
      	<td><a href="javascript:showLinkedItemList('<%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %>');" ><font color="red"><%= oPsItemList.FItemList(i).FPlusSaleItemCount %></font></a></td>
    </tr>
    <% next %>
    
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if oPsItemList.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPsItemList.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>
    
    		<% for i=0 + oPsItemList.StarScrollPage to oPsItemList.FScrollCount + oPsItemList.StarScrollPage - 1 %>
    			<% if i>oPsItemList.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>
    
    		<% if oPsItemList.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<p>
<!--
*상품검색 후, 팝업창에서 추가구성상품 등록<br>
*<br>
-->

<%
set oPsItemList = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
