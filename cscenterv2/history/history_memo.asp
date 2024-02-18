<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/history/cs_memocls.asp" -->
<%

dim i, userid, orderserial, phoneNumer
userid      = requestCheckVar(request("userid"),32)
orderserial = requestCheckVar(request("orderserial"),32)
phoneNumer  = requestCheckVar(request("phoneNumer"),32)

'==============================================================================
dim ocsmemo
set ocsmemo = New CCSMemo

ocsmemo.FRectUserID = userid
ocsmemo.FRectOrderserial = orderserial
ocsmemo.FRectPhoneNumber = phoneNumer

if (userid <> "") or (orderserial<>"") or (phoneNumer<>"") then
    ocsmemo.GetCSMemoList
end if

%>
<script>
function GotoHistoryMemoMidify(divcd,id,userid,orderserial) {
    var popwin = window.open("/cscenterv2/history/history_memo_write.asp?divcd="+divcd+"&id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial,"GotoHistoryMemoMidify","width=600 height=400 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>

<table width="100%" border=0 cellspacing=1 cellpadding=1 class=a bgcolor="F4F4F4">
<% if ocsmemo.FResultCount > 0 then %>
    <tr align="center" bgcolor="F3F3FF">
        <td height="20" width="50">구분</td>
    	<td width="50">idx</td>
     	<td width="80">고객ID</td>
    	<td width="80">주문번호</td>
    	<td>내용</td>
        <td width="80">등록자</td>
    	<td width="80">접수일</td>
    	<td width="30">완료</td>
    </tr>
    <tr>
        <td height="1" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>

<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="FFFFFF">
        <td height="20" >
     	  	 <%= ocsmemo.FItemList(i).GetDivCDName %>
        </td>
    	<td>
    		<%= ocsmemo.FItemList(i).Fid %>
    	</td>
     	<td><%= ocsmemo.FItemList(i).Fuserid %></td>
    	<td><%= ocsmemo.FItemList(i).Forderserial %></td>
    	<td align="left"><a href="javascript:GotoHistoryMemoMidify('<%= ocsmemo.FItemList(i).fdivcd %>','<%= ocsmemo.FItemList(i).Fid %>','<%= ocsmemo.FItemList(i).Fuserid %>','<%= ocsmemo.FItemList(i).Forderserial %>')">
    <acronym title="<%= ocsmemo.FItemList(i).Fcontents_jupsu %>"><%= DDotFormat(ocsmemo.FItemList(i).Fcontents_jupsu,25) %></acronym></a></td>
        <td><%= ocsmemo.FItemList(i).Fwriteuser %></td>
    	<td><acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><% if (ocsmemo.FItemList(i).Ffinishyn = "Y") then %>완료<% end if %></td>
    </tr>
    <tr>
        <td height="1" colspan="8" bgcolor="#CCCCCC"></td>
    </tr>
<% next %>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="6" align="center">검색결과가 없습니다.</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->