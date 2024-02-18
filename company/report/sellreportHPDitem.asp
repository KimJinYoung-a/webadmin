<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<%
'###############################################
' PageName : itemselllist.asp
' Discription : 업체어드민 상품판매 목록
' History : 2008.07.01 허진원 : 없는 날짜입력을 재계산 하도록 수정
'###############################################

dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim itemid, itemname
dim datetype
dim oldjumun


itemid = request("itemid")
itemname = request("itemname")
searchtype = request("searchtype")
searchrect = request("searchrect")
datetype   = request("datetype")
oldjumun = request("oldjumun")

if (datetype="") then datetype="jumunil"

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

'상품코드 유효성 검사(2008.07.11;허진원)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim cknodate,ckdelsearch
cknodate = request("cknodate")
ckdelsearch = request("ckdelsearch")

dim page
dim ojumun

page = request("page")
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

if ckdelsearch<>"on" then
	ojumun.FRectDelNoSearch="on"
end if


if searchtype="01" then
	ojumun.FRectBuyname = searchrect
elseif searchtype="02" then
	ojumun.FRectReqName = searchrect
elseif searchtype="03" then
	ojumun.FRectUserID = searchrect
elseif searchtype="04" then
	ojumun.FRectIpkumName = searchrect
elseif searchtype="06" then
	ojumun.FRectSubTotalPrice = searchrect
end if

ojumun.FRectItemid = itemid
ojumun.FRectItemName = itemname
ojumun.FRectDesignerID = "haepumdal"
ojumun.FPageSize = 100
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectOldJumun = oldjumun
ojumun.SearchJumunListByDesignerSelllist_HPDCase

dim ix,iy


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
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
			검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >주문일
			<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >결제일
			<input type="radio" name="datetype" value="beasong" <% if (datetype="beasong") then response.write "checked" %> >배송일
			(* 최근 6개월 이내 주문내역만 검색됩니다.)
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	검색결과 : 총 <b><font color="red"><% = ojumun.FTotalCount %></font></b> 개
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80" align="center">상품코드</td>
		<td>상품</td>
		<td>옵션</td>
		<td width="80">판매가</td>
		<td width="50">수량</td>
		<td width="100">판매가합계</td>
	</tr>
<% if ojumun.FResultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="12">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice, totalCNT
sumprice = ojumun.FMasterItemList(ix).FItemCost * ojumun.FMasterItemList(ix).FItemNo
%>
	
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a" align="center" bgcolor="#FFFFFF">
	<% else %>
	<tr class="gray" align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><%= ojumun.FMasterItemList(ix).FItemID  %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
		<td>&nbsp;</td>
		<% else %>
		<td><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
		<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(sumprice,0) %></td>
	</tr>
	 <%
	  totalsumprice =  totalsumprice + sumprice 
	  totalCNT = totalCNT + ojumun.FMasterItemList(ix).FItemNo
	 %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="4" height="25" >
		<td align="center"><% =FormatNumber(totalCNT,0) %></td>
		<td align="right"><b><% =FormatNumber(totalsumprice,0) %></b></td>
		
	</tr>
<% end if %>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ojumun = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->