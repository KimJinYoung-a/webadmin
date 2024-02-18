<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->

<%

dim i, userid

userid = request("userid")

'==============================================================================
dim ocsmileage
set ocsmileage = New CCSCenterMileage

ocsmileage.FRectUserID = userid

ocsmileage.GetCSCenterMileageSummary


'==============================================================================
dim ocsmileagelist
set ocsmileagelist = New CCSCenterMileage

ocsmileagelist.FRectUserID = userid

ocsmileagelist.GetCSCenterMileageList

%>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
body {
    background-color: #FFFFFF;
}

.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}
</style>
<table width="100%" border="0" cellspacing="0" cellpadding="2" class="a" bgcolor="FFFFFF">
    <tr>
        <td colspan="10" height="25">
			<% if (ocsmileage.FResultCount > 0) then %>
			구매마일리지[<b><%= FormatNumber(CLng(ocsmileage.FItemList(0).Ftotalbuymileage) + CLng(ocsmileage.FItemList(0).Ftotaloldbuymileage),0) %></b>] +
			보너스마일리지[<b><%= FormatNumber(ocsmileage.FItemList(0).Ftotalbonusmileage,0) %></b>] -
			사용마일리지[<b><%= FormatNumber(ocsmileage.FItemList(0).Ftotalspendmileage,0) %></b>] =
			잔여마일리지[<b><%= FormatNumber(CLng(ocsmileage.FItemList(0).Ftotalbuymileage) + CLng(ocsmileage.FItemList(0).Ftotaloldbuymileage) + CLng(ocsmileage.FItemList(0).Ftotalbonusmileage) - CLng(ocsmileage.FItemList(0).Ftotalspendmileage),0) %></b>]
			<% else %>
			검색결과가 없습니다.(탈퇴고객일 수 있습니다.)
			<% end if %>
        </td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
    <tr height="20" align="center" bgcolor="F3F3FF">
    	<td width="50">IDX</td>
      	<td width="60">마일리지</td>
      	<td width="50">구분</td>
      	<td width="70">&nbsp;&nbsp;적요코드</td>
      	<td>적요내용</td>
      	<td width="80">등록일</td>
      	<td width="90">관련주문번호</td>
      	<td width="60">삭제여부</td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
<% if (ocsmileagelist.FResultCount > 0) then %>
    <% for i = 0 to (ocsmileagelist.FResultCount - 1) %>
    <tr align="center" height="20" <% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td><%= ocsmileagelist.FItemList(i).Fid %></td>
    	<td align="right"><%= FormatNumber(ocsmileagelist.FItemList(i).Fmileage,0) %></td>
    	<td>
    	    <% if ocsmileagelist.FItemList(i).Fmileage >= 0 then %><font color="blue">적립</font><% else %><font color="red">사용</font><% end if %>
    	</td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyocd %></td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyo %></td>
    	<td><acronym title="<%= ocsmileagelist.FItemList(i).Fregdate %>"><%= Left(ocsmileagelist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsmileagelist.FItemList(i).Forderserial %></td>
    	<td><% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>삭제<% end if %></td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>
    <% next %>

<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="9">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>


<%

set ocsmileage = Nothing
set ocsmileagelist = Nothing

%>


<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
