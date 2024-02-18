<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고 확인 페이지
' History : 2007.09.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop_dailystock.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<%
const C_STOCK_DAY=7
dim itemgubun, itemid, itemoption 		'변수지정
itemgubun = request("itemgubun")		'상품구분을 받아온다
itemid = request("itemid")				'상품id 받아옴
itemoption = request("itemoption")		'상품옵션코드 받아옴
	if itemgubun="" then 				'상품구분이 공백이라면
		itemgubun="10"					'기본값 10 입력
	end if
	if itemoption="" then 				'상품옵션코드가 공백이면
		itemoption="0000"				'기본값 0000 입력
	end if

dim oitem
set oitem = new CItemInfo				'변수에 클래스 넣고
oitem.FRectItemID = itemid				'상품id를 넣고
	if itemid<>"" then					'상품id가 공백이면
		oitem.GetOneItemInfo
	end if

dim oitemoption							'아이템옵션부분
set oitemoption = new CItemOption		'클래스 넣고
oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

	if (oitemoption.FResultCount<1) then	'상품옵션코드가 1보다 작다면
		itemoption = "0000"					'기본값 0000 넣고
	end if

dim offstock			'오프라인재고파악
set offstock = new COffShopDailyStock		'클래스넣고
offstock.FRectItemGubun = itemgubun
offstock.FRectItemid = itemid
offstock.FRectItemoption = itemoption
	if itemid<>"" then
			if oitem.FResultCount>0 then
				offstock.FRectMakerid = oitem.FOneItem.FMakerid
			end if
		offstock.GetCurrentAllShopItemStock
	end if

dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummarystock										'온라인재고파악
set osummarystock = new CSummaryItemStock				'클래스 넣고
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
	if itemid<>"" then
		osummarystock.GetCurrentItemStock
		osummarystock.GetDaily_Logisstock_Summary
	end if

dim i,menupos
	menupos = request("menupos")
%>

<script language="javascript">

	function addfrm() {
		jaegoaddfrm.target= "view";
		jaegoaddfrm.submit();
	}

</script>

<!-- 표 검색부분 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>재고 확인</strong> / 동일상품은 등록되지 않습니다.</font>
			</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<p align="right"><a href="/admin/auction/auction_add_re.asp">일괄등록</a></p>
			상품코드: <input type=text name=itemid value="<%= itemid %>" size=9 maxlength=9>
			&nbsp;&nbsp;
			<input type=button value="검색" onclick="document.frm.submit();">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</form>
<form name="jaegoaddfrm" method=post action="/admin/auction/auction_process.asp">
<input type="hidden" name="fmode" value="item_add">
</table>
<!-- 표 검색부분 끝-->

<!-- 상품 정보 시작-->
<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">상품구분</td>
      	<td width="300">
      		<% if itemgubun = 10 then %>
				온라인상품
			<% elseif itemgubun = 90 then %>
				오프라인상품
			<% elseif itemgubun = 70 then %>
				소모품
			<% end if %>
      	</td>
      	<td width="60">배송구분 :</td>
      	<td colspan=2><%= oitem.FOneItem.GetDeliveryName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품코드 :</td>
      	<td><%= Format00(5,oitem.FOneItem.FItemID) %></td>
      	<td>전시여부 : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FDispyn) %>"><%= oitem.FOneItem.FDispyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID :</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>판매여부 : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명 :</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>사용여부 : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>

    <% if oitemoption.FResultCount>1 then %>

		<!-- 옵션이 있는경우 -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
		    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
		      	<td><font color="#AAAAAA">옵션명 :</font></td>
		      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).foptionname %></font></td>
		      	<td><font color="#AAAAAA">한정여부 : </font></td>
		      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>

		    <% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
		    <tr bgcolor="#EEEEEE">
		    <% else %>
		    <tr bgcolor="#FFFFFF">
		    <% end if %>
		      	<td>옵션명 :</td>
		      	<td><%= oitemoption.FITemList(i).foptionname %></td>
		      	<td>한정여부 : </td>
		      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo%></b>)</td>
		    </tr>
		    <% end if %>
	    <% next %>
    <% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>옵션코드 :</td>
	      	<td>-<input type="hidden" value="0000" name="itemoption"></td>
	      	<td>한정여부 : </td>
	      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>
</table>
<!-- 상품 정보 끝-->
</table>
<% dim oip
	set oip = new Cauctionlist        	'클래스 지정
	oip.Frectitemid = itemid
	oip.fwritelist()					'클래스를 실행
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
<td>옥션 카테고리 :</td>
<td><input type="text" name="auction_cate_code" value="10060500"></td>
<td>원산지 :</td>
<td><input type="text" name="wonsanji" value="한국"> ex) 한국 , 국외</td>
<td>
<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>,">
<input type="hidden" name="makerid" value="<%= oitem.FOneItem.FMakerid %>">
<input type="hidden" name="imagesrc" value="<%= oitem.FOneItem.FListImage %>">
<input type="hidden" name="itemname" value="<%= oitem.FOneItem.FItemName %>">
<input type="button" value="저장" onclick=addfrm();></td>
</tr></form>
<tr bgcolor="#FFFFFF">
<td>상품설명 :</td>
<td colspan="4"><textarea name="ten_itemcontent" cols="80" rows="30"><%= oip.flist(0).fitemcontent %></textarea></td>
</tr>
</table>

<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
    </tr>
</table>
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<%
set oitemoption = Nothing
set oitem = Nothing
set osummarystock = Nothing
%>

<iframe frameboarder=0 height=0 width=0 name="view" id="view"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->