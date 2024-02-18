<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->

<%

dim i, userid, showall, research
userid      = request("userid")
showall     = request("showall")
research    = request("research")

if (research="") and (showall="") then showall="on"

'==============================================================================
''현재 마일리지 합계
dim ocsmileage
set ocsmileage = New CCSCenterMileage
ocsmileage.FRectUserID = userid

if (ocsmileage.FRectUserID<>"") then
    ocsmileage.getUserCurrentMileage
end if

'==============================================================================
''마일리지 Log
dim ocsmileagelist
set ocsmileagelist = New CCSCenterMileage
if (showall<>"on") then
    ocsmileagelist.FRectDeleteYn = "N"
end if
ocsmileagelist.FRectUserID = userid

if (ocsmileagelist.FRectUserID<>"") then
    ocsmileagelist.GetCSCenterMileageList
end if

'==============================================================================
''만료예정  마일리지 합계
dim CExpireDT 
CExpireDT = Left(CStr(now()),4) + "-12-31"

dim oExpireMileTotal
set oExpireMileTotal = new CCSCenterMileage
oExpireMileTotal.FRectUserid = userid
oExpireMileTotal.FRectExpireDate = CExpireDT
if (userid<>"") then
    oExpireMileTotal.getNextExpireMileageSum
end if


%>
<script language='javascript'>
function popYearExpireMileList(yyyymmdd,userid){
    var popwin = window.open('popAdminExpireMileSummary.asp?yyyymmdd=' + yyyymmdd + '&userid=' + userid,'popAdminExpireMileSummary','width=660,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			아이디 : <input type="text" class="text" name="userid" value="<%= userid %>">
          	&nbsp;
          	<input type="checkbox" name="showall" <%= chkIIF(showall="on","checked","") %> >삭제내역도표시
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button" value="검색" onclick="document.frm.submit()">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>요약정보</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="85">6개월이전</td>
    	<td width="85">최근6개월</td>
    	<td width="85">아카데미(강좌)</td>
    	<td width="85">구매마일리지</td>
    	<td width="85">보너스마일리지</td>
    	<td width="85">사용마일리지</td>
    	<td width="85">소멸된마일리지</td>
      	<td width="85">잔여마일리지</td>
      	<td width="110">소멸예정 마일리지<br>(<%= oExpireMileTotal.FRectExpireDate %>)</td>
      	<td>비고</td>
    </tr>
<% if (ocsmileagelist.FResultCount > 0) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= FormatNumber(ocsmileage.FOneItem.Fflowerjumunmileage,0) %></td>
    	<td><%= FormatNumber(ocsmileage.FOneItem.Fjumunmileage,0) %></td>
    	<td><%= FormatNumber(ocsmileage.FOneItem.Facademymileage,0) %></td>
    	<td><%= FormatNumber(ocsmileage.FOneItem.getTotalBuymileage,0) %></td>
      	<td><font color="blue"><%= FormatNumber(ocsmileage.FOneItem.Fbonusmileage,0) %></font></td>
      	<td><font color="red"><%= FormatNumber(ocsmileage.FOneItem.Fspendmileage*(-1),0) %></font></td>
      	<td><font color="red"><%= FormatNumber(ocsmileage.FOneItem.FrealExpiredMileage*(-1),0) %></font></td>
      	<td><%= FormatNumber(ocsmileage.FOneItem.getCurrentMileage,0) %></td>
      	<td><a href="javascript:popYearExpireMileList('<%= oExpireMileTotal.FOneItem.FExpireDate %>','<%= userid %>');"><%= FormatNumber(oExpireMileTotal.FOneItem.getMayExpireTotal,0) %></a></td>
      	<td></td>
    </tr>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>0</td>
    	<td>0</td>
    	<td>0</td>
      	<td>0</td>
      	<td>0</td>
      	<td>0</td>
      	<td>0</td>
      	<td>0</td>
      	<td></td>
      	<td></td>
    </tr>
<% end if %>
</table>


<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>보너스마일리지 상세내역</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="120">아이디</td>
    	<td width="60">Idx</td>
      	<td width="60">마일리지</td>
      	<td width="50">구분</td>
      	<td width="80">적요코드</td>
      	<td width="200">적요</td>
      	<td width="80">등록일</td>
      	<td width="90">주문번호</td>
      	<td width="60">삭제여부</td>
      	<td>비고</td>
    </tr>
<% if (ocsmileagelist.FResultCount > 0) then %>
        <% for i = 0 to (ocsmileagelist.FResultCount - 1) %>
    <tr align="center" <% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td><%= ocsmileagelist.FItemList(i).Fuserid %></td>
    	<td><%= ocsmileagelist.FItemList(i).Fid %></td>
    	<td align="right">
    	    <% if ocsmileagelist.FItemList(i).Fmileage >= 0 then %><font color="blue"><%= ocsmileagelist.FItemList(i).Fmileage %></font><% else %><font color="red"><%= ocsmileagelist.FItemList(i).Fmileage %></font><% end if %>
    	</td>
    	<td>
    	    <% if ocsmileagelist.FItemList(i).Fmileage >= 0 then %><font color="blue">적립</font><% else %><font color="red">사용</font><% end if %>
    	</td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyocd %></td>
    	<td align="left"><%= ocsmileagelist.FItemList(i).Fjukyo %></td>
    	<td><acronym title="<%= ocsmileagelist.FItemList(i).Fregdate %>"><%= Left(ocsmileagelist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsmileagelist.FItemList(i).Forderserial %></td>
    	<td><% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>삭제<% end if %></td>
      	<td></td>
    </tr>
        <% next %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="10"> 검색된 내용이 없습니다.</td>
    </tr>
<% end if %>
</table>


<%

set ocsmileage = Nothing
set ocsmileagelist = Nothing
set oExpireMileTotal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->