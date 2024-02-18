

<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim nowdate,searchnextdate
dim designer,dateback, ipkumdiv

nowdate = Left(CStr(now()),10)

ipkumdiv = request("ipkumdiv")
designer = request("designer")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")



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

if designer<>"" then
ojumun.FRectDesignerID = designer
end if

ojumun.FPageSize = 50
ojumun.FCurrPage = page
ojumun.FRectIpkumdiv = ipkumdiv
ojumun.DesignerDateMiBaljuDetail

dim ix,iy
%>
<script language='javascript'>

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = '_blank';
    frm.action="/admin/ordermaster/viewordermaster.asp"
	frm.submit();

}

function ViewItem(itemid){
window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");
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
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<%	drawSelectBoxDesignerWithName "designer", designer %>
			검색기간(주문통보일) :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			주문상태:
			<select class="select" name="ipkumdiv">
		     	<option value="4" <% if ipkumdiv="4" then response.write "selected" %>>결제완료이상</option>
		     	<option value="5" <% if ipkumdiv="5" then response.write "selected" %>>주문통보이상</option>
	     	</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><% = ojumun.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80"">주문번호</td>
		<td width="100"">브랜드ID</td>
		<td width="60">주문자</td>
		<td width="60">수령인</td>
		<td>상품명<font color="blue">[옵션]</font></td>
		<td width="40">수량</td>
		<td width="40">취소<br>구분</td>
		<td width="65">입금확인일</td>
		<td width="65">발주일<br>(통보일)</td>
		<td width="50">발주후<br>경과일</td>
		<td width="60">주문확인</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%= ix %>" method="post" >
	<input type="hidden" name="orderserial" value="<%= ojumun.FMasterItemList(ix).FOrderSerial %>">
	<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
	<tr align="center" height="25" class="a" bgcolor="#FFFFFF">
	<% else %>
	<tr align="center" height="25" class="gray" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ix %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FOrderSerial %></a></td>
		<td><%= ojumun.FMasterItemList(ix).Fmakerid %></td>
		<td><%= ojumun.FMasterItemList(ix).FBuyname %></td>
		<td><%= ojumun.FMasterItemList(ix).FReqname %></td>
		<td align="left">
		    <a href="javascript:ViewItem(<% =ojumun.FMasterItemList(ix).FItemid  %>)"><%= ojumun.FMasterItemList(ix).FItemname %></a>
    		<% if (ojumun.FMasterItemList(ix).FItemoption="") then %>
    		<% else %>
    			<br>
    			<font color="blue">[<%= ojumun.FMasterItemList(ix).FItemoption %>]</font>
    		<% end if %>
		</td>
		<td align="center"><%= ojumun.FMasterItemList(ix).FItemcnt %></td>
		<td align="center">
		<% if ojumun.FMasterItemList(ix).FCancelYn <> "Y" then %>
		&nbsp;
		<% else %>
		<font color="red">주문취소</font>
		<% end if %>
		</td>
		<td align="center"><%= FormatDateTime(ojumun.FMasterItemList(ix).Fipkumdate,2) %></td>
		<td align="center">
    		<% if Not IsNULL(ojumun.FMasterItemList(ix).Fbaljudate) then %>
    		<%= FormatDateTime(ojumun.FMasterItemList(ix).Fbaljudate,2) %>
    		<% end if %>
		</td>
		<td align="center">D + <%= ojumun.FMasterItemList(ix).GetBaljuPassedDate %></td>
		<td align="center">
		<% if IsNULL(ojumun.FMasterItemList(ix).FCurrstate) or (ojumun.FMasterItemList(ix).FCurrstate=0) then %>
		통보이전
		<% elseif ojumun.FMasterItemList(ix).FCurrstate = 2 then %>
		<font color="red">업체통보</font>
		<% elseif ojumun.FMasterItemList(ix).FCurrstate = 3 then %>
		<font color="blue">주문확인</font>
		<% elseif ojumun.FMasterItemList(ix).FCurrstate = 7 then %>
		<font color="#339900">출고완료</font>
		<% end if %>
		</td>
	</tr>
	</form>
	<% next %>
<% end if %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" height="30" align="center">
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
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->