<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사은품 등록
' Hieditor : 2013.01.15 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer,page,usingyn , research, mageview, imageview, itemgubun, itemid, itemname
dim cdl, cdm, cds, i, PriceDiffExists , IsDirectIpchulContractExistsBrand ,publicbarcode
	designer    = RequestCheckVar(request("designer"),32)
	page        = RequestCheckVar(request("page"),9)
	usingyn     = RequestCheckVar(request("usingyn"),1)
	research    = RequestCheckVar(request("research"),9)
	imageview   = RequestCheckVar(request("imageview"),9)
	itemgubun   = RequestCheckVar(request("itemgubun"),16)
	itemid      = RequestCheckVar(request("itemid"),9)
	itemname    = RequestCheckVar(request("itemname"),32)
	publicbarcode    = RequestCheckVar(request("publicbarcode"),20)
	cdl         = RequestCheckVar(request("cdl"),3)
	cdm         = RequestCheckVar(request("cdm"),3)
	cds         = RequestCheckVar(request("cds"),3)
	if page="" then page=1
	if research<>"on" then usingyn="Y"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer
	ioffitem.FRectOnlyUsing = usingyn
	ioffitem.FRectItemgubun = itemgubun
	ioffitem.FRectItemID = itemid
	ioffitem.FRectItemName = html2db(itemname)
	ioffitem.FRectCDL = cdl
	ioffitem.FRectCDM = cdm
	ioffitem.FRectCDS = cds
	ioffitem.FRectpublicbarcode = publicbarcode

	ioffitem.GetOffNOnLineGiftItemList
%>

<script language='javascript'>

//수정
function pop_itemedit_gift_edit(ibarcode){
	var pop_itemedit_gift_edit = window.open('/admin/offshop/pop_itemedit_gift_edit.asp?barcode=' + ibarcode,'pop_itemedit_gift_edit','width=1024,height=600,resizable=yes,scrollbars=yes');
	pop_itemedit_gift_edit.focus();
}

//등록
function pop_itemedit_gift_new(){
	var pop_itemedit_gift_new;

	pop_itemedit_gift_new = window.open('/admin/offshop/pop_itemedit_gift_edit.asp','pop_itemedit_gift_new','width=1024,height=600,scrollbars=yes,resizable=yes');
	pop_itemedit_gift_new.focus();
}

function ReSearch(page){
	if(frm.itemid.value!=''){
		if (!IsDouble(frm.itemid.value)){
			alert('상품번호는 숫자만 가능합니다.');
			frm.itemid.focus();
			return;
		}
	}

	frm.page.value = page;
	frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.submit();
}

function jsSelectThisAndCloseWin(itemgubun, itemid, itemoption) {
	opener.ReActWithThis(itemgubun, itemid, itemoption);
	opener.focus();
	window.close();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 :
		<% drawSelectBoxDesignerwithName "designer",designer  %>
		&nbsp;
		구분:
		<input type="radio" name="itemgubun" value="" <% if itemgubun = "" then response.write " checked" %>> 전체
		<input type="radio" name="itemgubun" value="85" <% if itemgubun = "85" then response.write " checked" %>> ON사은품
		<input type="radio" name="itemgubun" value="80" <% if itemgubun = "80" then response.write " checked" %>> OFF사은품
		&nbsp;
     	사용:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="ReSearch('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="9" style="IME-MODE: disabled" />
		&nbsp;
		상품명 : <input type="text" class="text" name="itemname" value="<%= itemname %>" size="24" maxlength="32">
		&nbsp;
		범용바코드 : <input type="text" class="text" name="publicbarcode" value="<%= publicbarcode %>" size="20" maxlength="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >이미지보기
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" value="사은품 등록" onclick="pop_itemedit_gift_new()">
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
	<td colspan="25">
		검색결과 : <b><%= ioffitem.FTotalcount %></b>
		<% if ioffitem.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>

		<b><%= page %> / <%= ioffitem.FTotalpage %></b>

		<% if (ioffitem.FTotalpage - ioffitem.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (imageview<>"") then %>
	<td width="50">이미지</td>
	<% end if %>
	<td>브랜드ID</td>
	<td width="90">상품코드</td>
	<td>상품명</td>
	<td>옵션명</td>

	<td width="60">소비자가</td>
	<td width="60">판매가</td>

	<td width="60">매입가</td>
	<td width="60">매장<br>공급가</td>
	<td width="30">센터<br>매입<br>구분</td>

	<td width="30">사용<br>여부</td>

	<td width="50">비고</td>
</tr>
<% for i=0 to ioffitem.FresultCount -1 %>
<% if ioffitem.FItemlist(i).Fisusing="N" then %>
<tr bgcolor="#EEEEEE">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
	<% if (imageview<>"") then %>
		<td width="50" height="55">
			<img src="<%= ioffitem.FItemlist(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0>
		</td>
	<% end if %>
	<td height="30"><%= ioffitem.FItemlist(i).FMakerID %></td>
	<td align="center" >
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).Fitemgubun %>-<%= CHKIIF(ioffitem.FItemlist(i).Fshopitemid>=1000000,Format00(8,ioffitem.FItemlist(i).Fshopitemid),Format00(6,ioffitem.FItemlist(i).Fshopitemid)) %>-<%= ioffitem.FItemlist(i).Fitemoption %>
		</a>
	</td>
	<td>
		<a href="javascript:pop_itemedit_gift_edit('<%= ioffitem.FItemlist(i).GetBarCode %>')" onfocus="this.blur()">
		<%= ioffitem.FItemlist(i).FShopItemName %>
		</a>
	</td>
	<td>
		<%= ioffitem.FItemlist(i).FShopitemOptionname %>
		<% if ioffitem.FItemlist(i).FOnlineOptaddprice<>0 then %>
		    <br>옵션추가금액: <%= FormatNumber(ioffitem.FItemlist(i).FOnlineOptaddprice,0) %>
		<% end if %>
	</td>
    <td align="right" >
        <%= FormatNumber(ioffitem.FItemlist(i).FShopItemOrgprice, 0) %>
    </td>
	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).FShopItemprice, 0) %>
	</td>

	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).Fshopsuplycash, 0) %>
	</td>
	<td align="right" >
		<%= FormatNumber(ioffitem.FItemlist(i).Fshopbuyprice, 0) %>
	</td>
    <td align="center" ><%= ioffitem.FItemlist(i).FCenterMwDiv %></td>
	<td align="center" >
		<%= ioffitem.FItemlist(i).Fisusing %>
	</td>
	<td align="center" >
		<input type="button" class="button" value="선택" onclick="jsSelectThisAndCloseWin('<%= ioffitem.FItemlist(i).Fitemgubun %>', '<%= ioffitem.FItemlist(i).Fshopitemid %>', '<%= ioffitem.FItemlist(i).Fitemoption %>')">
	</td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="center">
	<% if ioffitem.HasPreScroll then %>
		<a href="javascript:ReSearch('<%= ioffitem.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
		<% if i>ioffitem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:ReSearch('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if ioffitem.HasNextScroll then %>
		<a href="javascript:ReSearch('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>

<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->