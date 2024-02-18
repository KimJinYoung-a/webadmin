<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdate,searchnextdate
dim orderserial,itemid,ojumun
dim topn,designer,page
dim ckpointsearch,ckipkumdiv4
dim ix,iy,cknodate
dim order_desum
dim rectdispy, rectselly
dim rectorderby
dim oldlist
Dim cdl,cdm,cds

yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
dd1 = requestCheckvar(request("dd1"),10)
yyyy2 = requestCheckvar(request("yyyy2"),10)
mm2 = requestCheckvar(request("mm2"),10)
dd2 = requestCheckvar(request("dd2"),10)
designer = requestCheckvar(request("designer"),32)
orderserial = requestCheckvar(request("orderserial"),32)
itemid = requestCheckvar(request("itemid"),10)
topn = requestCheckvar(request("topn"),10)
ckpointsearch = requestCheckvar(request("ckpointsearch"),10)
cknodate = requestCheckvar(request("cknodate"),10)
order_desum = requestCheckvar(request("order_desum"),10)
rectdispy = requestCheckvar(request("rectdispy"),10)
rectselly = requestCheckvar(request("rectselly"),10)
rectorderby = requestCheckvar(request("rectorderby"),10)
oldlist = requestCheckvar(request("oldlist"),10)

cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)

if rectorderby="" then rectorderby="cnttotal"

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

topn = request("topn")
if (topn="") then topn=100

set ojumun = new CCategoryReport

if cknodate="" then
	ojumun.FRectFromDate = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectToDate = searchnextdate
end If
ojumun.FPageSize = topn
ojumun.FRectCD1 = cdl
ojumun.FRectCD2 = cdm
ojumun.FRectCD3 = cds
ojumun.FRectOldJumun = oldlist

ojumun.CategorySearchBestsellerList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
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
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역<br>
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		검색갯수 :
		<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
        &nbsp;
        관리<!-- #include virtual="/common/module/categoryselectbox.asp"-->

		</td>
		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="7" height="25" align="right">검색결과 : 총 <font color="red"><% = ojumun.FResultCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr >
	<td width="100" align="center">상품번호</td>
	<td  align="center">상품</td>
	<td width="100" align="center">디자이너ID</td>
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
sumprice = ojumun.FItemList(ix).FItemCost * ojumun.FItemList(ix).FItemNo
%>
	<tr class="a">
		<td align="center" height="25"><a href="<%= wwwUrl %>/shopping/category_prd.asp?itemid=<%= ojumun.FItemList(ix).FItemID %>" class="zzz" target="_blank"><%= ojumun.FItemList(ix).FItemID  %></a></td>
		<td align="center"><%= ojumun.FItemList(ix).FItemName %></td>
		<td align="center"><%= ojumun.FItemList(ix).FMakerid %></td>
		<% if (ojumun.FItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= ojumun.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="center"><%= FormatNumber(ojumun.FItemList(ix).FItemCost,0)  %></td>
		<td align="center"><%= ojumun.FItemList(ix).FItemNo %></td>
		<td align="center"><%= FormatNumber(sumprice,0) %></td>
	</tr>
	 <% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr>
		<td colspan="7" height="25" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->