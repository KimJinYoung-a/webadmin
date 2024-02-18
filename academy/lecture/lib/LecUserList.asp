<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->
<%
dim searchfield, userid, orderserial, username, userhp, etcfield, etcstring, itemid
dim checkYYYYMMDD, checkJumunDiv, checkJumunSite
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim jumundiv, jumunsite


itemid = RequestCheckvar(request("itemid"),10)


dim page
dim ojumun

page = RequestCheckvar(request("page"),10)
if (page="") then page=1

set ojumun = new CLectureFingerOrder

ojumun.FPageSize = 200
ojumun.FCurrPage = page
ojumun.FRectItemID = itemid
ojumun.GetFingerRealOrderListByItemID


dim ix,i
dim totalavailcount


dim olecture
set olecture = new CLecture
olecture.FRectIdx = itemid
olecture.GetOneLecture



dim olecschedule
set olecschedule = new CLectureSchedule
olecschedule.FRectidx = itemid

olecschedule.GetOneLecSchedule



%>

<script language='javascript'>

function GotoOrderDetail(orderserial) {
        var popwin = window.open('/academy/lecture/lec_orderdetail.asp?orderserial=' + orderserial,'LecOrderDetail','width=800,height=400,scrollbars=yes,resizable=yes');
        popwin.focus();
}
window.print();
</script>
	<!-- 강좌 설명 -->
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="#000000">

		<tr bgcolor="#FFFFFF">
			<td width="110" align="center" bgcolor="#DDDDFF">강좌명</td>
			<td><%= olecture.FOneItem.Flec_title %></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="110" align="center" bgcolor="#DDDDFF">강사명</td>
			<td><%= olecture.FOneItem.Flecturer_id %> (<%= olecture.FOneItem.Flecturer_name %>)</td>
		</tr>
		<tr  bgcolor="#FFFFFF">
			<td width="110" align="center" bgcolor="#DDDDFF">강좌비</td>
			<td><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %></td>
		<tr>
			<td width="110" align="center" bgcolor="#DDDDFF">재료비</td>
			<td bgcolor="#FFFFFF" >
				<% if olecture.FOneItem.Fmatinclude_yn="Y" then %>
					포함(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
				<% else %>
					별도(<%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %>)
				<% end if %>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td width="110" align="center" bgcolor="#DDDDFF">강의횟수 및 시간</td>
			<td bgcolor="#FFFFFF">
				<%= olecture.FOneItem.Flec_period %><%= olecture.FOneItem.Flec_count %>회 &nbsp;&nbsp;&nbsp;<%= olecture.FOneItem.Flec_time %>시간
			</td>
		</tr>
		<tr bgcolor="#FFFFFF" >
			<td width="110" align="center" bgcolor="#DDDDFF" rowspan="<%= olecschedule.FResultCount  %>">강좌일시</td>
			<td bgcolor="#FFFFFF" colspan="2">
				<%= olecture.FOneItem.Flec_startday1 %> ~ <%= olecture.FOneItem.Flec_endday1 %>
				<% if (olecture.FOneItem.Flec_startday1<>olecschedule.FItemList(0).Fstartdate) or (olecture.FOneItem.Flec_endday1<>olecschedule.FItemList(0).Fenddate) then %>
					<br><b><%= olecschedule.FItemList(0).Fstartdate %> ~ <%= olecschedule.FItemList(0).Fenddate %></b>
				<% end if %>
			</td>
		</tr>
	</table>
<br>
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#000000">
	    <tr align="center" bgcolor="#DDDDFF">

	    	<td width="70">주문번호</td>
	    	<td width="50">거래상태</td>
	    	<td width="60">성명</td>
	    	<td width="90">UserID</td>
	    	<td width="40">수량</td>
	    	<td>기타</td>
	    </tr>
	    <% if ojumun.FresultCount<1 then %>
	    <tr bgcolor="#FFFFFF">
	    	<td colspan="20" align="center">[검색결과가 없습니다.]</td>
	    </tr>
	    <% else %>

		<% for ix=0 to ojumun.FresultCount-1 %>

		<% if ojumun.FItemList(ix).IsAvailJumun then %>
		<% totalavailcount = totalavailcount + ojumun.FItemList(ix).FItemNo %>
		<tr align="center" bgcolor="#FFFFFF" class="a">
		<% else %>
		<tr align="center" bgcolor="#EEEEEE" class="gray">
		<% end if %>
			<td><a href="javascript:GotoOrderDetail('<%= ojumun.FItemList(ix).FOrderSerial %>');"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
			<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
			<td><%= ojumun.FItemList(ix).Fentryname %></td>
			<td align="left"><font color="<%= ojumun.FItemList(ix).GetUserLevelColor %>"><%= ojumun.FItemList(ix).FUserID %></font></td>
			<td><%= ojumun.FItemList(ix).FItemNo %></td>
			<td>기타</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td colspan="4"></td>
			<td align="center"><%= totalavailcount %></td>
			<td colspan="5"></td>
		</tr>
	</table>
	<% end if %>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<%
set olecture = Nothing
set olecschedule = Nothing
set ojumun = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->