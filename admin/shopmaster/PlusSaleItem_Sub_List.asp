<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/PlusSaleItemCls.asp"-->

<%
dim page
dim makerid, itemidArr, itemname
dim cdl, cdm, cds
dim openstate, research, sellyn, mwdiv

page        = RequestCheckVar(request("page"),9)
makerid     = RequestCheckVar(request("makerid"),32)
itemidArr   = RequestCheckVar(request("itemidArr"),1024)
itemname    = RequestCheckVar(request("itemname"),64)
cdl         = RequestCheckVar(request("cdl"),3)
cdm         = RequestCheckVar(request("cdm"),3)
cds         = RequestCheckVar(request("cds"),3)
openstate   = RequestCheckVar(request("openstate"),32)
research    = RequestCheckVar(request("research"),32)
sellyn      = RequestCheckVar(request("sellyn"),9)
mwdiv      = RequestCheckVar(request("mwdiv"),9)


if (research="") and (openstate="") then openstate="openscheduled"

if (page="") then page=1
itemidArr = Trim(itemidArr)
itemname  = Trim(itemname)
if (Right(itemidArr,1)=",") then itemidArr = Left(itemidArr,Len(itemidArr)-1)

dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FCurrPage = Page
oPlusSaleItem.FRectMakerid  = makerid
oPlusSaleItem.FRectCDL      = cdl
oPlusSaleItem.FRectCDM      = cdm
oPlusSaleItem.FRectCDS      = cds
oPlusSaleItem.FRectItemIDArr= itemidArr
oPlusSaleItem.FRectItemName = itemname
oPlusSaleItem.FRectOpenState= openstate
oPlusSaleItem.FRectMwDiv    = mwdiv
oPlusSaleItem.FRectSellYn   = sellyn

oPlusSaleItem.GetPlusSaleSubItemList

dim i
%>

<script language='javascript'>

function PlusSaleItem_Edit(){
	var popwin = window.open('PlusSaleItem_Edit.asp','PlusSaleItem_Edit','width=900,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PlusSaleItem_Sub_Edit(iitemid){
	var popwin = window.open('PlusSaleItem_Sub_New.asp?itemid=' + iitemid,'PlusSaleItem_Sub_Edit','width=600,height=450,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PlusSaleItem_Sub_New(){
	var popwin = window.open('PlusSaleItem_Sub_New.asp','PlusSaleItem_Sub_New','width=600,height=450,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function showLinkedItemList(iitemid){
    var popwin = window.open('PlusSaleItem_Edit.asp?itemid=' + iitemid,'PlusSaleItem_Edit','');
    popwin.focus();
}

function NextPage(v){
	document.frm.page.value=v;
	document.frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
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
		    판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
			거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
			&nbsp;
			진행상태 :
			<select class="select" name="openstate">
              <option value="">전체</option>
              <option value="open" <%= ChkIIF(openstate="open","selected","") %> >진행중</option>
              <option value="scheduled" <%= ChkIIF(openstate="scheduled","selected","") %> >진행예정</option>
              <option value="openscheduled" <%= ChkIIF(openstate="openscheduled","selected","") %> >진행중+진행예정</option>
              <option value="expired" <%= ChkIIF(openstate="expired","selected","") %> >기간종료</option>
            </select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->



<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!-- <input type="button" class="button" value="신규등록1" onClick="PlusSaleItem_Edit();"> -->
			<input type="button" class="button" value="신규등록" onClick="PlusSaleItem_Sub_New();">
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
			검색결과 : <b><%= oPlusSaleItem.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= oPlusSaleItem.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="60">상품코드</td>
    	<td width="50">이미지</td>
      	<td width="100">브랜드ID</td>
      	<td>상품명</td>
      	<td width="60">현재<br>판매가</td>
		<td width="60">현재<br>매입가</td>
		<td width="40">현재<br>마진</td>
		<td width="30">계약<br>구분</td>
		<td width="120">시작일<br>종료일</td>
      	<td width="50">플러스<br>할인구분</td>
      	<td width="50">플러스<br>할인율</td>
      	<td width="50">플러스<br>할인가</td>
      	<td width="60">플러스<br>할인매입가</td>
      	<td width="50">플러스<br>할인마진</td>
      	<td width="50">진행상태</td>
      	<td width="50">메인<br>링크<br>상품수</td>
    </tr>
    <% if (oPlusSaleItem.FResultCount<1) then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="16" align="center">[검색 결과가 없습니다.]</td>
    </tr>
    <% else %>
    <% for i=0 to oPlusSaleItem.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="javascript:PlusSaleItem_Sub_Edit('<%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %>')"><%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %></a></td>
    	<td><img src="<%= oPlusSaleItem.FItemList(i).FImageSmall %>" width="50"></td>
      	<td><%= oPlusSaleItem.FItemList(i).FMakerID %></td>
      	<td><a href="javascript:PlusSaleItem_Sub_Edit('<%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %>')"><%= oPlusSaleItem.FItemList(i).FItemName %></a></td>
      	<td>

      	    <%= FormatNumber(oPlusSaleItem.FItemList(i).FOrgPrice,0) %>
      	    <% if oPlusSaleItem.FItemList(i).IsCurrentSaleItem then %>
      	    <br><font color=#F08050>(할)<%= FormatNumber(oPlusSaleItem.FItemList(i).FSellCash,0) %></font>
      	    <% end if %>
      	    <% if oPlusSaleItem.FItemList(i).IsCouponItem then %>
      	        <br><font color=#5080F0>(쿠)<%= FormatNumber(oPlusSaleItem.FItemList(i).GetCouponAssignPrice,0) %></font>
      	    <% end if %>
      	</td>
      	<td>

      	    <%= FormatNumber(oPlusSaleItem.FItemList(i).FOrgSuplycash,0) %>
      	    <% if oPlusSaleItem.FItemList(i).IsCurrentSaleItem then %>
      	    <br><font color=#F08050>(할)<%= FormatNumber(oPlusSaleItem.FItemList(i).FBuyCash,0) %></font>
      	    <% end if %>
      	</td>
      	<td><%= fnPercent(oPlusSaleItem.FItemList(i).FBuyCash,oPlusSaleItem.FItemList(i).FSellCash,1) %></td>
      	<td><%= fnColor(oPlusSaleItem.FItemList(i).FMwDiv,"mw") %></td>
      	<% if (oPlusSaleItem.FItemList(i).IsAlwaysTerms) then %>
      	<td align="center"> 상시진행 </td>
      	<% else %>
      	<td align="center"><%= oPlusSaleItem.FItemList(i).FPlusSaleStartDate %>
      	<br>
      	<%= Left(oPlusSaleItem.FItemList(i).FPlusSaleEndDate,10) %></td>
      	<% end if %>
      	<td><%= oPlusSaleItem.FItemList(i).getMaginFlagName %><br><%= oPlusSaleItem.FItemList(i).FPlusSaleMargin %>%</td>
      	<td><%= oPlusSaleItem.FItemList(i).FPlusSalePro %>%</td>
      	<td><%= FormatNumber(oPlusSaleItem.FItemList(i).getPlusSalePrice,0) %></td>
      	<td><%= FormatNumber(oPlusSaleItem.FItemList(i).getPlusSaleBuycash,0) %></td>
      	<td><%= fnPercent(oPlusSaleItem.FItemList(i).getPlusSaleBuycash,oPlusSaleItem.FItemList(i).getPlusSalePrice,1) %></td>
      	<td><font color="<%= oPlusSaleItem.FItemList(i).getCurrstateColor %>"><%= oPlusSaleItem.FItemList(i).getCurrstateName %></font></td>
      	<td><a href="javascript:showLinkedItemList('<%= oPlusSaleItem.FItemList(i).FPlusSaleItemID %>');"><%= oPlusSaleItem.FItemList(i).FLinkedItemCount %></a></td>
    </tr>
    <% next %>
    <% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if oPlusSaleItem.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPlusSaleItem.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oPlusSaleItem.StarScrollPage to oPlusSaleItem.FScrollCount + oPlusSaleItem.StarScrollPage - 1 %>
    			<% if i>oPlusSaleItem.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oPlusSaleItem.HasNextScroll then %>
    			<a href="javascript:NextPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>

<p>

*시작일과 종료일이 없을경우, 상시적용<br>
*날짜로 검색하여 예정/진행중/종료 구분(종료 먹표시)<br>


<%
set oPlusSaleItem = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
