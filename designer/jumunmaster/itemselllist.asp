<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
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
dim orderserial,itemid, itemname
dim datetype
dim oldjumun


orderserial = requestCheckVar(request("orderserial"), 32)
itemid = requestCheckVar(request("itemid"), 32)
itemname = requestCheckVar(request("itemname"), 128)
searchtype = requestCheckVar(request("searchtype"), 32)
searchrect = requestCheckVar(request("searchrect"), 32)
datetype   = requestCheckVar(request("datetype"), 32)
oldjumun = requestCheckVar(request("oldjumun"), 32)

if (datetype="") then datetype="jumunil"

yyyy1 = requestCheckVar(request("yyyy1"), 32)
mm1 = requestCheckVar(request("mm1"), 32)
dd1 = requestCheckVar(request("dd1"), 32)
yyyy2 = requestCheckVar(request("yyyy2"), 32)
mm2 = requestCheckVar(request("mm2"), 32)
dd2 = requestCheckVar(request("dd2"), 32)

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
cknodate = requestCheckVar(request("cknodate"), 32)
ckdelsearch = requestCheckVar(request("ckdelsearch"), 32)

dim page
dim ojumun

page = requestCheckVar(request("page"), 32)
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
ojumun.FRectDesignerID = session("ssBctID")
ojumun.FPageSize = 100
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.FRectDateType = datetype
ojumun.FRectOldJumun = oldjumun
ojumun.SearchJumunListByDesignerSelllist

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

<form name="frm" method="get" action="" style="margin:0px;">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<!-- 표 상단바 시작-->
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   		<tr height="10" valign="bottom">
			<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			<td background="/images/tbl_blue_round_02.gif"></td>
			<td background="/images/tbl_blue_round_02.gif"></td>
			<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td>
        		상품번호 : <input type="text" name="itemid" value="<%= itemid %>" size="11" maxlength="16">
				&nbsp;
				<!--
					 상품명 : <input type="text" name="itemname" value="<%= itemname %>" size="11">
				   -->
				<br>
				검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="radio" name="datetype" value="jumunil" <% if (datetype="jumunil") then response.write "checked" %> >주문일
				<input type="radio" name="datetype" value="ipkumil" <% if (datetype="ipkumil") then response.write "checked" %> >결제일

				(* 최근 6개월 이내 주문내역만 검색됩니다.)
			</td>
			<td align="right">
        		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
			</td>
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
	</table>
	<!-- 표 상단바 끝-->

</form>

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
		<td width="80">공급가</td>
		<td width="50">수량</td>
		<td width="100">공급가합계</td>
	</tr>
	<% if ojumun.FResultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="12">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% for ix=0 to ojumun.FResultCount -1 %>
	<%
	Dim sumprice,totalsumprice
	sumprice = ojumun.FMasterItemList(ix).FBuycash * ojumun.FMasterItemList(ix).FItemNo
	%>
	<form name="frmOnerder_<%= ojumun.FMasterItemList(ix).FOrderSerial %>" method="post" >

		<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
		<tr class="a" align="center" bgcolor="#FFFFFF">
			<% else %>
			<tr class="gray" align="center" bgcolor="#FFFFFF">
				<% end if %>
				<td>
					<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
					<input type="hidden" name="menupos" value="<%= menupos %>">
					<input type="hidden" name="sitename" value="<%= ojumun.FMasterItemList(ix).FSiteName %>">
					<input type="hidden" name="userid" value="<%= ojumun.FMasterItemList(ix).FUserID %>">
					<%= ojumun.FMasterItemList(ix).FItemID  %>
				</td>
				<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
				<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
				<td>&nbsp;</td>
				<% else %>
				<td><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
				<% end if %>
				<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FItemCost,0)  %></td>
				<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FBuycash,0)  %></td>
				<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
				<td align="right"><%= FormatNumber(sumprice,0) %></td>
			</tr>
	</form>
	<% totalsumprice =  totalsumprice + sumprice %>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="10" height="25" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font> 원&nbsp;&nbsp;</td>
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
	<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
	<!-- #include virtual="/lib/db/dbclose.asp" -->
