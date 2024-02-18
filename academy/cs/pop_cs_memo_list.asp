<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%

dim i, userid, orderserial, searchfield, searchstring, finishyn, sitegubun

userid = RequestCheckvar(request("userid"),32)
orderserial = RequestCheckvar(request("orderserial"),16)
searchfield = RequestCheckvar(request("searchfield"),16)
searchstring = rRequestCheckvar(equest("searchstring"),128)

if (searchstring = "") then
        searchfield = ""
end if

finishyn = request("finishyn")
if finishyn="" then finishyn="A"

sitegubun = request("sitegubun")
if sitegubun="" then sitegubun="academy"


'==============================================================================
dim ocsmemo
set ocsmemo = New CCSMemo

if (searchfield = "userid") then
        userid = searchstring
        ocsmemo.FRectUserID = userid
elseif (searchfield = "orderserial") then
        orderserial = searchstring
        ocsmemo.FRectOrderserial = orderserial
end if

if (finishyn = "N") then
        ocsmemo.FRectIsFinished = "N"
end if

ocsmemo.FRectSiteGubun = sitegubun
ocsmemo.GetCSMemoList

%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
		    <img src="/images/icon_star.gif" align="absbottom">
		    <font color="red"><strong>CS�޸� ����</strong></font>
		</td>
		<td align="right" background="/images/tbl_blue_round_06.gif">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
   	<form name="frm" method="get" action="">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
            <select name="searchfield">
                <option value="" <% if (searchfield = "") then %>selected<% end if %>>����</option>
                <option value="userid" <% if (searchfield = "userid") then %>selected<% end if %>>���̵�</option>
                <option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>�ֹ���ȣ</option>
            </select>
            &nbsp;
            <input type="text" name="searchstring" value="<%= searchstring %>" size="12">&nbsp;&nbsp;
            &nbsp;
            <input type="radio" name="finishyn" value="A" <% if (finishyn = "A") then response.write "checked" end if %>>��ü
            <input type="radio" name="finishyn" value="N" <% if (finishyn = "N") then response.write "checked" end if %>>��ó����û�޸�
            &nbsp;&nbsp;
            |
            &nbsp;&nbsp;
            <input type="radio" name="sitegubun" value="all" <% if (sitegubun = "all") then response.write "checked" end if %>>��ü����Ʈ
            <input type="radio" name="sitegubun" value="academy" <% if (sitegubun = "academy") then response.write "checked" end if %>>�ΰŽ�
            <input type="radio" name="sitegubun" value="10x10" <% if (sitegubun = "10x10") then response.write "checked" end if %>>���ֹ�
        </td>
        <td align="right" valign="top"><input type="button" value="�˻�" onclick="document.frm.submit()"></td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr bgcolor="#DDDDFF">
        <td width="100" align="center">��ID</td>
        <td width="80" align="center">�ֹ���ȣ</td>
        <td width="60" align="center">����</td>
        <td width="50" align="center">������</td>
        <td align="center">����</td>
        <td width="70" align="center">������</td>
        <td width="60" align="center">ó������</td>
        <td width="90" align="center">ó����</td>
        <td width="70" align="center">ó����</td>
    </tr>
<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF">
        <td height="25"><%= ocsmemo.FItemList(i).Fuserid %></td>
        <td><%= ocsmemo.FItemList(i).Forderserial %></td>
        <td align="left"><%= ocsmemo.FItemList(i).GetDivCDName %></td>
        <td align="left"><%= ocsmemo.FItemList(i).Fwriteuser %></td>
        <td align="left"><%= Left(ocsmemo.FItemList(i).Fcontents_jupsu,40) %></td>
        <td align="center"><acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym></td>
        <td><%= ocsmemo.FItemList(i).Ffinishyn %></td>
        <td align="left"><%= ocsmemo.FItemList(i).Ffinishuser %></td>
        <td><acronym title="<%= ocsmemo.FItemList(i).Ffinishdate %>"><%= Left(ocsmemo.FItemList(i).Ffinishdate,10) %></acronym></td>
    </tr>
<% next %>
<% if (ocsmemo.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td height="25" colspan="9">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="center" align="right">
          &nbsp;
	</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<%

set ocsmemo = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->