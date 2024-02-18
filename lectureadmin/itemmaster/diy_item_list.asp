<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<%

dim itemid, makerid, itemname
dim sellyn, isusing, limityn, mwdiv
dim page

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)



if (sellyn="") then sellyn="A"

if (page="") then page=1

''if (isusing="") then isusing="Y"
''사용하는 상품만 표시로 변경
isusing="Y"

'상품코드 유효성 검사(2008.08.01;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname
oitem.FRectLimityn = limityn
oitem.FRectMWDiv = mwdiv
oitem.FPageSize = 30
oitem.FCurrPage = page


if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if


oitem.GetItemList

dim i

%>
<script>
function viewBySite(itemid){
    <% if (now()<#2016-09-06#) then %>
   // alert('9월 5일 사이트가 개편됩니다. 현재 보이는 페이지와 9월5일 이후 페이지가 다르니 이미지 수정은 9월 5일 이후 하시기 바랍니다.\r\n\r\n현재페이지 이미지타입 정사각형, \r\n개편된 페이지 상품이미지타입 직사각형');
    <% end if %>
    window.open('<%=wwwFingers%>/diyshop/shop_prd.asp?itemid='+itemid,'_blank'); 
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	SubmitSearch();
}
function SubmitSearch(){
	if ((document.frm.itemid.value != "") && ((document.frm.itemid.value*0) != 0)) {
	    alert("상품코드에는 숫자만 입력이 가능합니다.");
	    document.frm.itemid.focus();
	    return;
    }
	document.frm.submit();
}


// ============================================================================
// 기본정보수정
function editItemInfo(itemid) {
	var param = "itemid=" + itemid + "&fingerson=on";

	popwin = window.open('diy_item_infomodify.asp?' + param ,'editItemInfo','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// 옵션수정
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('diy_item_optionmodify.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function editSimpleItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/academy/comm/pop_diy_simpleitemedit.asp?' + param ,'editSimpleItemOption','width=500,height=650,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// 이미지수정
function editItemImage(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('diy_item_imagemodify.asp?' + param ,'editItemImage','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>


<!-- 표 상단바 시작-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상품코드 :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) SubmitSearch();">
			&nbsp;
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="20" onKeyPress="if (event.keyCode == 13) SubmitSearch();"><br>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:SubmitSearch();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
			&nbsp;
	     	한정:<% drawSelectBoxLimitYN "limityn", limityn %>
	     	&nbsp;
	     	거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
		</td>
	</tr>
	</form>
</table>

<p>

	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td width="60">상품코드</td>
			<td width="50">이미지</td>
			<td>상품명</td>
			<td width="30">거래<br>구분</td>
			<td width="30">판매<br>여부</td>
			<td width="40">한정<br>여부</td>
			<td width="60">판매가</td>
			<td width="60">공급가</td>
			<td width="50">기본<br>정보</td>
			<td width="50">이미지</td>
			<td width="70">한정판매<br>판매여부</td>
	    </tr>
<% if oitem.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
	    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
    <% for i=0 to oitem.FresultCount-1 %>
    	<% if (oitem.FItemList(i).Fisusing = "N") then %>
    	<tr class="a" height="25" bgcolor="<%= adminColor("gray") %>">
		<% else %>
		<tr class="a" height="25" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center"><%= oitem.FItemList(i).Fitemid %></td>
			<td align="center"><img src="<%= oitem.FItemList(i).Fsmallimage %>" width="50" height="50" border="0" alt=""></td>
			<% if (FALSE) then %>
			<td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%= oitem.FItemList(i).Fitemid %>" target="_blank"><font color="blue">(확인하기)</font></a></td>
		    <% else %>
		    <td align="left"><% =oitem.FItemList(i).Fitemname %>&nbsp;&nbsp;<a href="javascript:viewBySite('<%= oitem.FItemList(i).Fitemid %>');" ><font color="blue">(확인하기)</font></a></td>
	        <% end if %>
			<td align="center">
				<font color="<%= mwdivColor(oitem.FItemList(i).Fmwdiv) %>"><%= mwdivName(oitem.FItemList(i).Fmwdiv) %></font>
			</td>

			<td align="center"><%= fnColor(oitem.FItemList(i).Fsellyn,"yn") %></td>
			<td align="center">
        		<% if (oitem.FItemList(i).Flimityn = "Y") then %>
             		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
             		<br>(<%= (oitem.FItemList(i).Flimitno - oitem.FItemList(i).Flimitsold) %>)
        		<% else %>
              		<%= fnColor(oitem.FItemList(i).Flimityn,"yn") %>
       			<% end if %>
			</td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Fsellcash,0) %></td>
			<td align="right"><%= FormatNumber(oitem.FItemList(i).Fbuycash,0) %></td>
		    <td align="center">
		    	<a href="javascript:editItemInfo('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
		    	<a href="javascript:editItemImage('<%= oitem.FItemList(i).FItemId %>')">
		    	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		    	</a>
		    </td>
		    <td align="center">
        <% if (oitem.FItemList(i).Fmwdiv = "U") then %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<img src="/images/icon_modify.gif" border="0" align="absbottom">
		      	</a>
        <% else %>
		      	<a href="javascript:editSimpleItemOption('<%= oitem.FItemList(i).FItemId %>')">
		      	<b>[</b>수정요청<b>]</b>
		      	</a>
        <% end if %>

		    </td>
		</tr>
		<% next %>
	</table>
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->