<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim shopid, disp, sort, diff, isusing, mwdiv

shopid = request("shopid")
disp = trim(request("disp"))
sort = trim(request("sort"))
diff = trim(request("diff"))
isusing = request("isusing")
mwdiv = request("mwdiv")

if (disp = "") then
        disp = "availsysstock"
end if
if (sort = "") then
        sort = "makerid"
end if
if (diff = "") then
        diff = "-100"
end if
if (diff > 0) then
        diff = -1 * diff
end if
if (sort = "") then
        sort = "makerid"
end if
if ((request("research") = "") and (isusing = "")) then
	''isusing = "on"
end if


'==============================================================================
dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectMakerid = shopid

osummarystock.FRectKindDisplay = disp
osummarystock.FRectKindSort = sort
osummarystock.FRectParameter = diff
osummarystock.FRectOnlyIsUsing = isusing
'osummarystock.FRectStartDate = BasicMonth + "-01"

osummarystock.FRectMWDiv = mwdiv
osummarystock.FRectUseYN = isusing

osummarystock.GetCurrentStockByOnlineBrandMinus

dim i
%>

<script language='javascript'>

function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function CheckNSellDispYN(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.sellyn[0].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "Y";
					}else if (frm.sellyn[1].checked){
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "S";
					}else{
						upfrm.sellyn.value = upfrm.sellyn.value + "|" + "N";
					}
                    /*
					if (frm.dispyn[0].checked){
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "Y";
					}else{
						upfrm.dispyn.value = upfrm.dispyn.value + "|" + "N";
					}
					*/
				}
			}
		}
		frm.submit();
	}
}
</script>

<p>
<script>
function SubmitForm()
{
        document.frm.submit();
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	브랜드: <% drawSelectBoxDesignerwithName "shopid",shopid %>
			&nbsp;
			거래구분 :
			<select class="select" name="mwdiv">
				<option value="">-선택-</option>
				<option value="M" <% if (mwdiv = "M") then %>selected<% end if %> >매입</option>
				<option value="W" <% if (mwdiv = "W") then %>selected<% end if %> >위탁</option>
				<option value="U" <% if (mwdiv = "U") then %>selected<% end if %> >업체</option>
				<option value="Z" <% if (mwdiv = "Z") then %>selected<% end if %> >미지정</option>
			</select>
			&nbsp;
			상품 사용여부 :
			<select class="select" name="isusing">
				<option value="">-선택-</option>
				<option value="Y" <% if (isusing = "Y") then %>selected<% end if %> >사용함</option>
				<option value="N" <% if (isusing = "N") then %>selected<% end if %> >사용않함</option>
			</select>

        	<br>

            표시방식 :
            <input type="radio" name="disp" value="totsysstock" <% if (disp = "totsysstock") then %>checked<% end if %>> 시스템재고&nbsp;&nbsp;
        	<input type="radio" name="disp" value="availsysstock" <% if (disp = "availsysstock") then %>checked<% end if %>> 유효재고&nbsp;&nbsp;
        	<input type="radio" name="disp" value="realstock" <% if (disp = "realstock") then %>checked<% end if %>> 실사재고&nbsp;&nbsp;
        	<input type="radio" name="disp" value="diff" <% if (disp = "diff") then %>checked<% end if %>> 대량오차&nbsp;&nbsp;
        	범위 <input type="text" name="diff" value="<%= diff %>" size="6">

        	<br>

        	표시순서 :
        	<input type="radio" name="sort" value="makerid" <% if (sort = "makerid") then %>checked<% end if %>> 브랜드&nbsp;&nbsp;
        	<input type="radio" name="sort" value="itemid" <% if (sort = "itemid") then %>checked<% end if %>> 신상품&nbsp;&nbsp;
        	<input type="radio" name="sort" value="diff" <% if (sort = "diff") then %>checked<% end if %>> 오차&nbsp;&nbsp;
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a><br>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        <% if (disp = "availsysstock") then %>
	        	* 유효재고가 <%= diff %> 개보다 같거나 작은 상품에 대해, "<%= sort %>" 순으로 정렬합니다.
			<% end if %>
			<% if (disp = "realstock") then %>
			    * 실사재고가 <%= diff %> 개보다 같거나 작은 상품에 대해, "<%= sort %>" 순으로 정렬합니다.
			<% end if %>
			<% if (disp = "diff") then %>
			    * 오차가 <%= abs(diff) %> 개보다 같거나 큰 상품에 대해, "<%= sort %>" 순으로 정렬합니다.
			<% end if %>
        </td>
        <td align="right">검색결과 : <%= FormatNumber(osummarystock.FresultCount,0) %></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">이미지</td>
		<td width="80">브랜드</td>
		<td width="40">상품<br>코드</td>
		<td>상품명<br>(옵션명)</td>
		<td width="35">배송<br>구분</td>
        <td width="50">전체<br>입고<br>반품</td>
        <td width="50">전체<br>판매<br>반품</td>
        <td width="50">전체<br>출고<br>반품</td>
        <td width="50">기타<br>출고<br>반품</td>
        <td width="50">CS<br>출고<br>재고</td>
		<td width="50"><b>시스템<br>재고</b></td>

		<td width="50">총<br>실사<br>오차</td>
		<td width="50">실사<br>재고</td>
		<td width="50">총<br>불량</td>
		<td width="50">유효<br>재고</td>

		<td width="60">한정<br>여부</td>
		<!-- <td width="30">전시<br>여부</td> -->
		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="30">품절<br>여부</td>
    </tr>
<% for i=0 to osummarystock.FresultCount-1 %>
	<% if osummarystock.FItemList(i).Fisusing="Y" then %>
    <tr bgcolor="#FFFFFF" align=center>
    <% else %>
    <tr bgcolor="#EEEEEE" align=center>
    <% end if %>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
		<td align="left">
	          <%= osummarystock.FItemList(i).FMakerID %>
	    </td>
		<td>
	          <a href="javascript:PopItemSellEdit('<%= osummarystock.FItemList(i).FItemID %>');"><%= osummarystock.FItemList(i).FItemID %></a>
	    </td>
		<td align="left">
			<a href="javascript:PopItemDetail('<%= osummarystock.FItemList(i).FItemID %>','<%= osummarystock.FItemList(i).FItemOption %>')"><%= osummarystock.FItemList(i).FItemName %></a>
        	<% if (osummarystock.FItemList(i).FItemOptionName <> "") then %>
          	<br>(<%= osummarystock.FItemList(i).FItemOptionName %>)
        <% end if %>
        </td>
        <td><font color="<%= mwdivColor(osummarystock.FItemList(i).Fmwdiv) %>"><%= mwdivName(osummarystock.FItemList(i).Fmwdiv) %></font></td>
		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Ftotipgono, 0) %>&nbsp;</td>
		<td align="right"><%= FormatNumber(-1*osummarystock.FItemList(i).Ftotsellno, 0) %>&nbsp;</td>
		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Foffchulgono + osummarystock.FItemList(i).Foffrechulgono, 0) %>&nbsp;</td>
        <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono, 0) %>&nbsp;</td>
        <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Ferrcsno, 0) %>&nbsp;</td>
		<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ftotsysstock, 0) %></b>&nbsp;</td>

		<td align="right"><b><%= FormatNumber(osummarystock.FItemList(i).Ferrrealcheckno, 0) %></b>&nbsp;</td>
        <td align="right"><%= FormatNumber(osummarystock.FItemList(i).getErrAssignStock, 0) %>&nbsp;</td>
		<td align="right"><%= FormatNumber(osummarystock.FItemList(i).Ferrbaditemno, 0) %>&nbsp;</td>
        <td align="right"><%= FormatNumber(osummarystock.FItemList(i).Frealstock, 0) %>&nbsp;</td>

		<td>
        	<% if (osummarystock.FItemList(i).Flimityn = "Y") then %>
	          	한정(<%= osummarystock.FItemList(i).GetLimitStr %>)
	            <% if (osummarystock.FItemList(i).Foptlimityn = "Y") then %>
	            <br>(<%= osummarystock.FItemList(i).Foptlimitno %>/<%= osummarystock.FItemList(i).Foptlimitsold %>)
	            <% else %>
	            <br>(<%= osummarystock.FItemList(i).FLimitNo %>/<%= osummarystock.FItemList(i).FLimitSold %>)
	          	<% end if %>
        	<% end if %>
        </td>
        <!-- <td></td> -->
		<td><font color="<%= ynColor(osummarystock.FItemList(i).Fsellyn) %>"><%= osummarystock.FItemList(i).Fsellyn %></font></td>
		<td><font color="<%= ynColor(osummarystock.FItemList(i).Fisusing) %>"><%= osummarystock.FItemList(i).Fisusing %></font></td>
        <td><% if osummarystock.FItemList(i).IsSoldOut  then %><font color="red">품절</font><% end if %></td>
</tr>
<% next %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<form name="frmArrupdate" method="post" action="dolimitsoldset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="dispyn" value="">
<input type="hidden" name="sellyn" value="">
</form>
</table>
<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
