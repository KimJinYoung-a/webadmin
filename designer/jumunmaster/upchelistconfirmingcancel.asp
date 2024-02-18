<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim orderserial,dateback
dim searchType, searchValue

nowdate = Left(CStr(now()),10)

orderserial = requestCheckVar(request("orderserial"), 32)

yyyy1   = requestCheckVar(request("yyyy1"), 32)
mm1     = requestCheckVar(request("mm1"), 32)
dd1     = requestCheckVar(request("dd1"), 32)
yyyy2   = requestCheckVar(request("yyyy2"), 32)
mm2     = requestCheckVar(request("mm2"), 32)
dd2     = requestCheckVar(request("dd2"), 32)
searchType  = requestCheckVar(request("searchType"), 32)
searchValue  = requestCheckVar(request("searchValue"), 32)


if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy1,mm1, dd1 - 30)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

dim cknodate
cknodate = requestCheckVar(request("cknodate"), 32)

dim page
dim ojumun

page = requestCheckVar(request("page"), 32)
if (page="") then page=1

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FPageSize = 100
ojumun.FCurrPage = page
ojumun.FRectSearchType = SearchType
ojumun.FRectSearchValue = SearchValue
ojumun.FRectDesignerID = session("ssBctID")
ojumun.DesignerDateBaljuCancleList

dim ix,iy
%>
<script language='javascript'>

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="/designer/viewordermaster.asp"
	frm.submit();

}

function ViewItem(itemid){
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

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select class="select" name="searchType" >
				<option value="">검색조건</option>
				<option value="orderserial" <%= ChkIIF(searchType="orderserial","selected","") %> >주문번호</option>
				<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >상품코드</option>
				<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >구매자</option>
				<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >수령인</option>
			</select>
			<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13" maxlength="11">
			&nbsp;
			검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;(출고기준일<!--주문통보일-->)
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="FFFFFF">
		<td height="25" colspan="15">
			검색결과 : <b><% = ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalpage %></b>
		</td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">주문번호</td>
		<td width="50">주문자</td>
		<td width="50">수령인</td>
		<td>상품명<font color="blue">&nbsp;[옵션]</font></td>
		<td width="80">수량</td>
		<td width="100">취소구분</td>
		<td width="100">입금확인일</td>
		<td width="100">출고기준일<!--주문통보일--></td>
		<td width="120">주문확인</td>
	</tr>
	<% if ojumun.FresultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="10">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr align="center" class="a" bgcolor="#FFFFFF">
	<% else %>
	<tr align="center" class="gray" bgcolor="#FFFFFF">
	<% end if %>
		<td height="25"><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
		<td align="left">
			<a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
			<% if (ojumun.FMasterItemList(ix).FItemoption<>"") then %>
				<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
			<% end if %>
		</td>
		<td><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
		<td>
			<% if (ojumun.FMasterItemList(ix).FDetailCancelyn="Y") then %>
			<font color="red">상품취소</font>
			<% elseif (ojumun.FMasterItemList(ix).FCancelyn<>"N") then %>
			<font color="red">주문취소</font>
			<% else %>
			&nbsp;
			<% end if %>
		</td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).FIpkumdate %>"><%= left(ojumun.FMasterItemList(ix).FIpkumdate,10) %></acronym></td>
		<td><acronym title="<%= ojumun.FMasterItemList(ix).Fbaljudate %>"><%= left(ojumun.FMasterItemList(ix).Fbaljudate,10) %></acronym></td>
		<td>
			<% if ojumun.FMasterItemList(ix).FCurrstate = 0 then %>
			<font color="red">주문미확인</font>
			<% else %>
			<font color="blue">주문확인</font>
			<% end if %>
		</td>
	</tr>
	</form>
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
<!-- 표 하단바 끝-->


<%
set ojumun = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->