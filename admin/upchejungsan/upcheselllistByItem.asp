<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매출통계
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim designer,itemid,dateback
dim datetype,delivertype,vPurchaseType, sitename, isPlusSaleItem

nowdate = Left(CStr(now()),10)


designer = request("designer")
itemid = request("itemid")
datetype   = request("datetype")
delivertype = request("delivertype")
isPlusSaleItem = request("isPlusSaleItem")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
vPurchaseType = requestCheckVar(request("purchasetype"),2)
sitename = requestCheckVar(request("sitename"),32)


if (datetype="") then datetype="jumunil"
if (delivertype="") then delivertype="upche"

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

dateback = DateSerial(yyyy1,mm2, dd2-7)

yyyy1 = Left(dateback,4)
mm1   = Mid(dateback,6,2)
dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim cknodate
cknodate = request("cknodate")

dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CBaljuMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if


ojumun.FRectDesignerID = designer
ojumun.FRectItemid = itemid
ojumun.FPageSize = 100
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectDeliverType = delivertype
ojumun.FRectBrandPurchaseType = vPurchaseType
ojumun.FRectSiteName = sitename
ojumun.FRectIsPlusSaleItem = isPlusSaleItem

ojumun.DesignerDateSellListByItem

dim ix,iy

dim returnRate
%>
<script language='javascript'>

function ViewOrderDetail(iorderserial){
	var popwin;
    popwin = window.open('/admin/ordermaster/viewordermaster.asp?orderserial=' + iorderserial,'orderdetail','scrollbars=yes,resizable=yes,width=800,height=600');
    popwin.focus();
}

function popOrderDetailEdit(idx){
	var popwin = window.open('/common/orderdetailedit_UTF8.asp?idx=' + idx,'orderdetailedit','width=600,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");
}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>



<p>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesigner "designer", designer %>
			&nbsp;
			상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="6">
			&nbsp;
			검색기간 :
			<select class="select" name="datetype">
		     	<option value='jumunil' <% if (datetype = "jumunil") then %>selected<% end if %> >주문일기준</option>
		     	<option value='ipkumil' <% if (datetype = "ipkumil") then %>selected<% end if %> >결제일기준</option>
		     	<option value='chulgoil' <% if (datetype = "chulgoil") then %>selected<% end if %> >출고일기준</option>
				 <option value='jungsanil' <% if (datetype = "jungsanil") then %>selected<% end if %> >정산일기준</option>
	     	</select>
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	배송구분
			<select class="select" name="delivertype">
		     	<option value="all" <% if delivertype="all" then response.write "selected" %> >전체</option>
		     	<option value="ten" <% if (delivertype="ten") then response.write "selected" %> >자체배송</option>
		     	<option value="upche" <% if (delivertype="upche") then response.write "selected" %> >업체배송</option>
	     	</select>
	     	&nbsp;|&nbsp;
    		구매유형 : 
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchaseType,"" %>
			&nbsp;|&nbsp;
			추가구매 : 
			<% drawSelectBoxIsPlusSaleItem "isPlusSaleItem", isPlusSaleItem %>
    		&nbsp;|&nbsp;
    		사이트 :
    		<% 'drawSelectBoxOnIpjumShop "sitename",sitename %>
    		<% Drawsitename "sitename",sitename %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			총판매수량 : <%= FormatNumber(ojumun.FSumItemNo,0) %>
			&nbsp;
			총판매가 : <%= FormatNumber(ojumun.FSumItemCost,0) %>
			&nbsp;
			총매입가 : <%= FormatNumber(ojumun.FSumBuyCash,0) %>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>브랜드ID</td>
		<td width="50">상품코드</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="30">수량</td>
		<td width="65">소비자가<br>(총액)</td>
		<td width="65">판매총액<br>(할인적용)</td>
		<td width="65">구매총액<br>(쿠폰적용)</td>
		<td width="65">매입총액<br>(쿠폰적용)</td>
		<td width="65">매출수익<br>(구매-매입)</td>
		<td width="50">수익률</td>
		<td>비고</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
		<td colspan="15" align="center">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>

	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojumun.FMasterItemList(ix).FMakerid %></td>
		<td><a href="javascript:popOrderDetailEdit(<%= ojumun.FMasterItemList(ix).Fdetailidx %>);"><%= ojumun.FMasterItemList(ix).FItemid %></a></td>
		<td align="left">
			<a href="javascript:popOrderDetailEdit(<%= ojumun.FMasterItemList(ix).Fdetailidx %>);"><%= ojumun.FMasterItemList(ix).FItemname %></a>
			<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
			<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= ojumun.FMasterItemList(ix).FItemcnt %></font>
			<% else %>
			<%= ojumun.FMasterItemList(ix).FItemcnt %>
			<% end if %>
		</td>
		<td align="right">
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FOrgitemCost,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FOrgitemCost,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FitemcostCouponNotApplied,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FitemcostCouponNotApplied,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FSellCash,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FSellCash,0) %>
			<% end if %>
		</td>
		<td align="right">
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber(ojumun.FMasterItemList(ix).FBuyCash,0) %></font>
			<% else %>
			<%= FormatNumber(ojumun.FMasterItemList(ix).FBuyCash,0) %>
			<% end if %>
		</td>
		<td>
			<% if ojumun.FMasterItemList(ix).FJumunDiv="9" then %>
			<font color="red"><%= FormatNumber((ojumun.FMasterItemList(ix).FSellCash - ojumun.FMasterItemList(ix).FBuyCash),0) %></font>
			<% else %>
			<%= FormatNumber((ojumun.FMasterItemList(ix).FSellCash - ojumun.FMasterItemList(ix).FBuyCash),0) %>
			<% end if %>
		</td>
		<td>
			<%
			if (ojumun.FMasterItemList(ix).FSellCash = 0) then
				returnRate = 0
			else
				returnRate = (ojumun.FMasterItemList(ix).FSellCash - ojumun.FMasterItemList(ix).FBuyCash) / ojumun.FMasterItemList(ix).FSellCash * 100
			end if
			%>
			<%= FormatNumber(returnRate,2) %> %
		</td>
		<td>

		</td>
	</tr>
	<% next %>
<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			<% for ix=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if ix>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= ix %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
