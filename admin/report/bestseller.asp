<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.06.03 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate
dim orderserial,itemid,ojumun
dim topn,designer,page
dim ckpointsearch,ckipkumdiv4
dim ix,iy,cknodate
dim order_desum
dim rectdispy, rectselly
dim rectorderby
dim oldlist

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
designer = request("designer")
orderserial = request("orderserial")
itemid = request("itemid")
topn = request("topn")
ckpointsearch = request("ckpointsearch")
cknodate = request("cknodate")
order_desum = request("order_desum")
rectdispy = request("rectdispy")
rectselly = request("rectselly")
rectorderby = request("rectorderby")
oldlist = request("oldlist")

if rectorderby="" then rectorderby="cnttotal"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

topn = request("topn")
if (topn="") then topn=100

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FRectItemid = itemid
ojumun.FRectDesignerID = designer
ojumun.FPageSize = topn
ojumun.FRectckpointsearch = ckpointsearch
ojumun.FRectIpkumDiv4 = "on" 'ckipkumdiv4
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.FRectDispY = rectdispy
ojumun.FRectSellY = rectselly
ojumun.FRectOrderBy = rectorderby
ojumun.FRectOldJumun = oldlist

ojumun.SearchBestsellerList
%>
<script language='javascript'>
function ViewOrderDetail(itemid){


window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");


}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ReSearch(ifrm){
	var v = ifrm.topn.value;
	if (!IsDigit(v)){
		alert('숫자만 가능합니다.');
		ifrm.topn.focus();
		return;
	}

	if (v>1000){
		alert('천건 이하만 검색가능합니다.');
		ifrm.topn.focus();
		return;
	}
	ifrm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		디자이너 :
		<% drawSelectBoxDesigner "designer",designer %>
		검색갯수 :
		<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" ><br>
		<input type=radio name="rectorderby" value="cnttotal" <% if rectorderby="cnttotal" then response.write "checked" %> >건수
		<input type=radio name="rectorderby" value="sumtotal" <% if rectorderby="sumtotal" then response.write "checked" %> >금액
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		</td>
		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
	<td colspan="9" height="25" align="right">검색결과 : 총 <font color="red"><% = ojumun.FResultCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td width="60" align="center">상품번호</td>
	<td width="50" align="center">이미지</td>
	<td width="100" align="center">브랜드ID</td>
	<td  align="center">상품명</td>
	<td width="80" align="center">옵션</td>
	<td width="100" align="center">가격</td>
	<td width="65" align="center">총갯수</td>
	<td width="90" align="center">판매총액</td>
	<td width="90" align="center">수익총액</td>
</tr>
<% if ojumun.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice, buyprice
sumprice = ojumun.FMasterItemList(ix).FItemCost * ojumun.FMasterItemList(ix).FItemNo
buyprice = ojumun.FMasterItemList(ix).FBuycash * ojumun.FMasterItemList(ix).FItemNo
%>
	<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="25"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= ojumun.FMasterItemList(ix).FItemID %>" class="zzz" target="_blank"><%= ojumun.FMasterItemList(ix).FItemID  %></a></td>
		<td align="center"><img src="<%= ojumun.FMasterItemList(ix).Fsmallimage %>" width=50 height=50></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FMakerid %></td>
		<td ><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(sumprice,0) %></td>
		<td align="right"><%= FormatNumber(sumprice-buyprice,0) %></td>
	</tr>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr>
		<td colspan="9" height="25" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->