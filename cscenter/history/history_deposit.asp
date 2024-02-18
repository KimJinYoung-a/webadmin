<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_depositcls.asp" -->

<%

dim i, userid

userid = request("userid")

'==============================================================================
dim ocsdeposit
set ocsdeposit = New CCSCenterDeposit

ocsdeposit.FRectUserID = userid

ocsdeposit.GetCSCenterDepositSummary


'==============================================================================
dim ocsdepositlist
set ocsdepositlist = New CCSCenterDeposit

ocsdepositlist.FRectUserID = userid

ocsdepositlist.GetCSCenterDepositList

'' response.write "aaa"
'' response.end


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
        <td colspan="10">
			취득예치금[<b><%= FormatNumber(ocsdeposit.FOneItem.Fgaindeposit, 0) %></b>]
			-
			사용예치금[<b><%= FormatNumber(ocsdeposit.FOneItem.Fspenddeposit, 0) %></b>]
			=
			잔여예치금[<b><%= FormatNumber(ocsdeposit.FOneItem.Fcurrentdeposit, 0) %></b>]
        </td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
    <tr height="20" align="center" bgcolor="F3F3FF">
    	<td width="50">IDX</td>
      	<td width="60">예치금</td>
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
<% if (ocsdepositlist.FResultCount > 0) then %>
    <% for i = 0 to (ocsdepositlist.FResultCount - 1) %>
    <tr align="center" height="20" <% if (ocsdepositlist.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td><%= ocsdepositlist.FItemList(i).Fidx %></td>
    	<td align="right"><%= FormatNumber(ocsdepositlist.FItemList(i).Fdeposit,0) %></td>
    	<td>
    	    <% if ocsdepositlist.FItemList(i).Fdeposit >= 0 then %><font color="blue">적립</font><% else %><font color="red">사용</font><% end if %>
    	</td>
    	<td><%= ocsdepositlist.FItemList(i).Fjukyocd %></td>
    	<td><%= ocsdepositlist.FItemList(i).Fjukyo %></td>
    	<td><acronym title="<%= ocsdepositlist.FItemList(i).Fregdate %>"><%= Left(ocsdepositlist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsdepositlist.FItemList(i).Forderserial %></td>
    	<td><% if (ocsdepositlist.FItemList(i).Fdeleteyn = "Y") then %>삭제<% end if %></td>
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

set ocsdeposit = Nothing
set ocsdepositlist = Nothing

%>


<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
