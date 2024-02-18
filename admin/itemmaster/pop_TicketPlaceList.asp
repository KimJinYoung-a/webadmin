<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/ticketItemCls.asp"-->

<%
Dim txPlacename, page, sortMethod
txPlacename = requestCheckvar(request("txPlacename"),32)
page = requestCheckvar(request("page"),10)
sortMethod = requestCheckvar(request("srtMtd"),1)

if (page="") then page=1
if (sortMethod="") then sortMethod="n"


Dim oticketPlace
set oticketPlace = new CTicketPlace
oticketPlace.FPageSize=20
oticketPlace.FCurrPage = page
oticketPlace.FsortMethod = sortMethod
oticketPlace.FRectTicketPlaceName = txPlacename
oticketPlace.getTicketPlaceList

dim i
%>
<script language='javascript'>
function modiThis(iid){
    location.href="pop_TicketPlaceInfo.asp?ticketPlaceIdx="+iid;
}

function selThis(iidx,iname){
    opener.ticketreg.ticketPlaceIdx.value=iidx;
    opener.ticketreg.ticketPlaceName.value=iname;
    window.close();
    
}

function goPage(ipg){
	document.frm.page.value= ipg;
	document.frm.submit();
}

function chgSort(srt){
	document.frm.page.value= 1;
	document.frm.srtMtd.value= srt;
	document.frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" action="" method="get">
<input type="hidden" name="menupos" value="1106">
<input type="hidden" name="page" >
<input type="hidden" name="srtMtd" value="<%=sortMethod%>" >
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
		공연장소 명 <input type="text" name="txPlacename" value="<%= txPlacename %>" size="20" maxlength="50" class="text">
     	<input type="submit" value="검색" class="button" onfocus="this.blur();">
	</td>
</tr>
</form>
</table>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="margin:10px 0 10px 0;">
<tr>
	<td align="left">
		<input type="button" class="button" value="신규등록" onClick="location.href='pop_TicketPlaceInfo.asp?ticketPlaceIdx=0';">
	</td>
	<td align="right">
		정렬방법 :
		<select id="selSort" onchange="chgSort(this.value);" class="select">
			<option value="n" <%=chkIIF(sortMethod="n","selected","")%>>최신순</option>
			<option value="r" <%=chkIIF(sortMethod="r","selected","")%>>등록순</option>
		</select>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4">
			검색결과 : <b><%= oticketPlace.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> /<%= oticketPlace.FTotalPage %></b>
		</td>
	</tr>
	</form>
	<tr align="center" bgcolor="#E6E6E6">
		<td width="60">No.</td>
		<td width=200> 공연장소 명</td>
		<td width="200">주소</td>
		<td width="200">비고</td>
    </tr>
    <% for i=0 to oticketPlace.FResultCount-1 %>
	<tr class="a" height="25" bgcolor="#FFFFFF">
		<td><%= oticketPlace.FItemList(i).FticketPlaceIdx %></td>
		<td><%= oticketPlace.FItemList(i).FticketPlaceName %></td>
		<td><%= Left(oticketPlace.FItemList(i).FtPAddress,20) %></td>
		<td align="center">
		    <input type="button" value="수정" onClick="modiThis('<%= oticketPlace.FItemList(i).FticketPlaceIdx %>');">
		    &nbsp;&nbsp;
		    <input type="button" value="선택" onClick="selThis('<%= oticketPlace.FItemList(i).FticketPlaceIdx %>','<%= Replace(oticketPlace.FItemList(i).FticketPlaceName,"'","") %>');">
		</td>
	</tr>
	<% next %>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="4" align="center">
			<% if oticketPlace.HasPreScroll then %>
			<a href="javascript:goPage('<%= oticketPlace.StartScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oticketPlace.StartScrollPage to oticketPlace.FScrollCount + oticketPlace.StartScrollPage - 1 %>
    			<% if i>oticketPlace.FTotalpage then Exit for %>
    			<% if CStr(page)=CStr(i) then %>
    			<font color="red">[<%= i %>]</font>
    			<% else %>
    			<a href="javascript:goPage('<%= i %>')">[<%= i %>]</a>
    			<% end if %>
    		<% next %>

    		<% if oticketPlace.HasNextScroll then %>
    			<a href="javascript:goPage('<%= i %>')">[next]</a>
    		<% else %>
    			[next]
    		<% end if %>
		</td>
	</tr>
</table>
<%
set oticketPlace = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->