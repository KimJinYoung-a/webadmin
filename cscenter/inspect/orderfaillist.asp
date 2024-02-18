<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->

<%
dim page, research, notdel
dim yyyy1, mm1, yyyy2, mm2, dd1,dd2
dim stdate, nowdate,searchnextdate

nowdate = Left(CStr(now()),10)
stdate = Left(CStr(DateAdd("m",-1,now())),10)

page = requestCheckvar(request("page"),10)
if page="" then page=1

yyyy1   = request("yyyy1")
yyyy2   = request("yyyy2")
mm1     = request("mm1")
mm2     = request("mm2")
dd1     = request("dd1")
dd2     = request("dd2")
notdel  = request("notdel")
research= request("research")

if (research="") and (notdel="") then notdel="on"

if (yyyy1="") then
    yyyy1 = Left(stdate,4)
	mm1   = Mid(stdate,6,2)
	dd1   = Mid(stdate,9,2)

	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim ojumun
set ojumun = new CJumunMaster
ojumun.FCurrPage = page
ojumun.FPageSize = 20
ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ojumun.FRectRegEnd = searchnextdate
ojumun.FRectDelNoSearch = notdel
ojumun.FRectIpkumdiv = "0"
ojumun.SearchJumunList

dim i
%>
<script type='text/javascript'>
function NextPage(ipage){
	document.frm_search.page.value= ipage;
	document.frm_search.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="" onSubmit="return chk_form(this)">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			주문일 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	        &nbsp;
	        <input type="checkbox" name="notdel" <%= chkIIF(notdel="on","checked","") %> > 삭제건 제외
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ojumun.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ojumun.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="100">주문번호</td>
		<td align="center" width="100">UserID</td>
		<td align="center" width="100">지불수단</td>
		<td align="center">TID</td>
		<td align="center" width="100">결제상태</td>
		<td align="center" width="100">취소구분</td>
	</tr>
	<% for i=0 to ojumun.FResultCount -1 %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td><%= ojumun.FMasterItemList(i).FOrderSerial %></td>
	    <td><%= printUserId(ojumun.FMasterItemList(i).FUserId,2,"**") %></td>
	    <td><%= ojumun.FMasterItemList(i).JumunMethodName %></td>
	    <td><%= ojumun.FMasterItemList(i).Fpaygatetid %></td>
	    <td><font color="<%= ojumun.FMasterItemList(i).IpkumDivColor %>"><%= ojumun.FMasterItemList(i).IpkumDivName %></font></td>
	    <td><font color="<%= ojumun.FMasterItemList(i).CancelYnColor %>"><%= ojumun.FMasterItemList(i).CancelYnName %></font></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF" height="25">
	    <td colspan="6" align="center">
		    <% if ojumun.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ojumun.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
	
			<% for i=0 + ojumun.StartScrollPage to ojumun.FScrollCount + ojumun.StartScrollPage - 1 %>
				<% if i>ojumun.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
	
			<% if ojumun.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
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

