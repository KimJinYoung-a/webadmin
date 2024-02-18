<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.10.26 �̻� ����
'			2012.05.23 �ѿ�� �̵�/����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
dim i, userid, orderserial, phoneNumer, writeuser, finishyn
dim searchType, searchValue ,id
	userid      = requestCheckVar(request("userid"),32)
	orderserial = requestCheckVar(request("orderserial"),32)
	phoneNumer  = requestCheckVar(request("phoneNumer"),32)
	searchType = requestCheckVar(request("searchType"),32)
	searchValue = requestCheckVar(request("searchValue"),32)
	writeuser = requestCheckVar(request("writeuser"),32)
	finishyn  = requestCheckVar(request("finishyn"),32)
	id        = requestCheckVar(request("id"),32)

if (finishyn="") then finishyn="A"
if (searchType="PH") and (searchValue<>"") then phoneNumer=searchValue
if (searchType="UID") and (searchValue<>"") then userid=searchValue
if (searchType="OD") and (searchValue<>"") then orderserial=searchValue

if (phoneNumer<>"") then
    searchType = "PH"
    searchValue = phoneNumer
end if

if (userid<>"") then
    searchType = "UID"
    searchValue = userid
end if

if (orderserial<>"") then
    searchType = "OD"
    searchValue = orderserial
end if

dim ocsmemo
set ocsmemo = New CCSMemo
	
	''and �˻� ����.
	if (searchType="UID") then ocsmemo.FRectUserID = userid
	if (searchType="OD") then ocsmemo.FRectOrderserial = orderserial
	if (searchType="PH") then ocsmemo.FRectPhoneNumber = phoneNumer
	
	ocsmemo.FRectWriteUser = writeuser
	if (finishyn = "N") then
	    ocsmemo.FRectIsFinished = "N"
	end if
	
	if (userid <> "") or (orderserial<>"") or (phoneNumer<>"") or (writeuser<>"") or (finishyn="N")then
	    ocsmemo.GetCSMemoList
	end if

%>

<script type='text/javascript'>

//cs�޸� �Է�&����
function GotoHistoryMemo(id){

	<% if InStr(request.ServerVariables("HTTP_REFERER"),"popSiteReceive.asp")>0 or InStr(request.ServerVariables("HTTP_REFERER"),"iframeHistory.asp")>0 then %>
		var csmemoreg = window.open('/common/cscenter/cscenter_memo.asp?id='+id,'csmemoreg','width=800,height=450,scrollbars=yes,resizable=yes');
		csmemoreg.focus();
	<% end if %>
}

</script>

<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border=0 cellspacing=1 cellpadding=1 class=a bgcolor="EEEEEE">
<form name="frmSearch" method="get">
<tr height="20">
    <td>
        <select name="searchType" class="select">
	        <option value="OD" <%= chkIIF(searchType="OD","selected","") %> >�ֹ���ȣ
	        <option value="UID" <%= chkIIF(searchType="UID","selected","") %> >���̵�
	        <option value="PH" <%= chkIIF(searchType="PH","selected","") %> >��ȭ��ȣ
        </select>
        <input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="14">
        <!--
        ������:<input type="text" class="text" name="writeuser" value="<%= writeuser %>" size="10">
        -->
        <input type="radio" name="finishyn" value="A" <% if (finishyn = "A") then response.write "checked" end if %>>��ü
        <input type="radio" name="finishyn" value="N" <% if (finishyn = "N") then response.write "checked" end if %>>��ó��
        <input type="button" class="button" value="�˻�" onClick="frmSearch.submit()">
    </td>
</tr>
</form>
</table>

<table width="100%" border="0" cellpadding="2" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocsmemo.FResultCount > 0 then %>
    <tr>
        <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="30">����</td>
     	<td width="80">��ID<br><font color="blue">�ֹ���ȣ</font></td>
    	<td>����</td>
        <td width="65">�����<br><font color="blue">ó����</font></td>
    	<td width="65">������<br><font color="blue">�Ϸ���</font></td>
    	<td width="30">�Ϸ�<br>����</td>
    	<td width="45">���</td>
    </tr>
    <tr>
        <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>

<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="<%= chkIIF(CStr(ocsmemo.FItemList(i).Fid)=id,"#DDDDDD","#FFFFFF") %>">
        <td><%= ocsmemo.FItemList(i).GetDivCDName %><br><%= ocsmemo.FItemList(i).Fid %></td>
     	<td><%= ocsmemo.FItemList(i).Fuserid %><br><font color="blue"><%= ocsmemo.FItemList(i).Forderserial %></font></td>
    	<td align="left">
        	<% if (Trim(ocsmemo.FItemList(i).Fcontents_jupsu) = "") then %>
        		(�������)
        	<% else %>
        		<%= DDotFormat(ocsmemo.FItemList(i).Fcontents_jupsu,25) %>
        	<% end if %>
        </td>
        <td>
        	<%= ocsmemo.FItemList(i).Fwriteuser %>
        	<% if ocsmemo.FItemList(i).FDivCD<>"1" then %>
        	<br><font color="blue"><%= ocsmemo.FItemList(i).Ffinishuser %></font>
    		<% end if %>
        </td>
    	<td>
    		<acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym>
    		<% if ocsmemo.FItemList(i).FDivCD<>"1" then %>
    		<br><acronym title="<%= ocsmemo.FItemList(i).Ffinishdate %>"><font color="blue"><%= Left(ocsmemo.FItemList(i).Ffinishdate,10) %></font></acronym>
    		<% end if %>
    	</td>
    	<td><% if (ocsmemo.FItemList(i).Ffinishyn = "Y") then %>�Ϸ�<% end if %></td>
    	<td>
    		<input type="button" onclick="GotoHistoryMemo('<%= ocsmemo.FItemList(i).Fid %>');" value="��" class="button">
    	</td>
    </tr>
    <tr>
        <td height="1" colspan="8" bgcolor="#CCCCCC"></td>
    </tr>
<% next %>

<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="6" align="center">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->