<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="lib/classes/cscenter/cs_aslistcls.asp"-->
<%

dim i, userid, orderserial, searchtype

userid = request("userid")
orderserial = request("orderserial")
searchtype = request("searchtype")

if searchtype="" then searchtype="searchfield"

'==============================================================================
dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectSearchType = searchtype
if (userid <> "") then
        ocsaslist.FRectUserID = userid
else
        ocsaslist.FRectOrderSerial = orderserial
end if

ocsaslist.GetCSASMasterList

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
    <tr height="20" align="center" bgcolor="F3F3FF">
    	<td width="40">idx</td>
    	<td width="60">구분</td>
     	<td width="80">브랜드ID</td>
    	<td>제목</td>
        <td width="60">접수자</td>
    	<td width="65">접수일</td>
    	<td width="60">처리자</td>
    	<td width="65">처리일</td>
    	<td width="80">상태</td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
<% if (ocsaslist.FResultCount > 0) then %>
        <% for i = 0 to (ocsaslist.FResultCount - 1) %>
    <tr height="20" align="center" <% if (ocsaslist.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
    	<td><%= ocsaslist.FItemList(i).Fid %></td>
        <td><acronym title="<%= ocsaslist.FItemList(i).FdivcdName %>"><%= Left(ocsaslist.FItemList(i).FdivcdName,4) %></acronym></td>
     	<td><%= ocsaslist.FItemList(i).Fmakerid %></td>
    	<td align="left"><%= ocsaslist.FItemList(i).Ftitle %></td>
        <td><%= ocsaslist.FItemList(i).Fwriteuser %></td>
    	<td><acronym title="<%= ocsaslist.FItemList(i).Fregdate %>"><%= Left(ocsaslist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsaslist.FItemList(i).Ffinishuser %></td>
    	<td><acronym title="<%= ocsaslist.FItemList(i).Ffinishdate %>"><%= Left(ocsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
    	<td><%= ocsaslist.FItemList(i).Fcurrstatename %></td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>
        <% next %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="15">검색결과가 없습니다.</td>
    </tr>
<% end if %>
</table>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
