<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
사용안함 - 서팀 문의 요망
<%
dbget.close()	:	response.End
%>
<%
dim searchtype
dim searchrect

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim orderserial,itemid,designer
dim oldlist

nowdate = Left(CStr(now()),10)

designer = request("designer")
orderserial = request("orderserial")
itemid = request("itemid")
searchtype = request("searchtype")
searchrect = request("searchrect")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
oldlist = request("oldlist")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim cknodate,ckdelsearch,ckipkumdiv4
dim datetype
cknodate = request("cknodate")
ckdelsearch = request("ckdelsearch")
ckipkumdiv4 = request("ckipkumdiv4")
datetype = request("datetype")
if (datetype="") then datetype="jumunil"

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
ojumun.FRectDesignerID = designer
ojumun.FPageSize = 100
ojumun.FRectIpkumDiv4 = ckipkumdiv4
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectOldJumun = oldlist
ojumun.SearchJumunListByupcheSelllist

dim ix,iy

'response.write ojumun.FRectOrderSerial
'dbget.close()	:	response.End

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
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" ><input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역<br>
		디자이너 :
		<% drawSelectBoxDesigner "designer",designer %>&nbsp;
		item번호 :
		<input type="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16">
		&nbsp;<br>
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >주문일
		<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >결제일
		<input type="radio" name="datetype" value="beadal" <% if (datetype="beadal") then response.write "checked" %> >배송일
		(<input type="checkbox" name="ckipkumdiv4" <% if ckipkumdiv4="on" then response.write "checked" %> >결제완료이상검색)
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="7" height="25" align="right">검색결과 : 총 <font color="red"><% = ojumun.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr >
	<td width="100" align="center">상품번호</td>
	<td  align="center">상품</td>
	<td width="80" align="center">옵션</td>
	<td width="100" align="center">가격</td>
	<td width="65" align="center">총갯수</td>
	<td width="100" align="center">합계금액</td>
</tr>
<% if ojumun.FResultCount<1 then %>
<tr>
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
Dim sumprice,totalsumprice
sumprice = ojumun.FMasterItemList(ix).FItemCost * ojumun.FMasterItemList(ix).FItemNo
%>
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr class="a">
	<% else %>
	<tr class="gray">
	<% end if %>
		<td align="center" height="25"><%= ojumun.FMasterItemList(ix).FItemID  %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemName %></td>
		<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="center"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemNo %></td>
		<td align="center"><%= FormatNumber(sumprice,0) %></td>
	</tr>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr>
		<td colspan="7" height="25" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>

	<tr>
		<td colspan="7" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ojumun.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
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