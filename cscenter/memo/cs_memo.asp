<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.01.01 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/upchebeasongcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim i, userid, orderserial, searchfield, searchstring, finishyn, writeuser,qadiv, contents
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, dateback, MMGubunExvlude, mmgubun, retrydateexclude
Dim page
	userid          = requestCheckVar(request("userid"),32)
	orderserial     = requestCheckVar(request("orderserial"),32)
	searchfield     = requestCheckVar(request("searchfield"),32)
	searchstring    = requestCheckVar(request("searchstring"),32)
	contents    	= requestCheckVar(request("contents"),32)
	writeuser       = requestCheckVar(request("writeuser"),32)
	qadiv           = requestCheckVar(request("qadiv"),32)
	MMGubunExvlude  = request("MMGubunExvlude")	'requestCheckVar(,32)
	mmgubun  		= requestCheckVar(request("mmgubun"),32)
	retrydateexclude  		= requestCheckVar(request("retrydateexclude"),32)
	finishyn = requestCheckVar(request("finishyn"),32)
	yyyy1   = requestCheckVar(request("yyyy1"),4)
	mm1     = requestCheckVar(request("mm1"),2)
	dd1     = requestCheckVar(request("dd1"),2)
	yyyy2   = requestCheckVar(request("yyyy2"),4)
	mm2     = requestCheckVar(request("mm2"),2)
	dd2     = requestCheckVar(request("dd2"),2)
	page     = requestCheckVar(request("page"),10)

If page = "" Then page = 1

if (searchstring = "") then
	searchfield = ""
end if
if (finishyn="") then finishyn="A"

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)

	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1

    dateback = DateSerial(yyyy2,mm2-2, dd2)

    yyyy1 = Left(dateback,4)
    mm1   = Mid(dateback,6,2)
    dd1   = Mid(dateback,9,2)

    MMGubunExvlude = "on"
end if

dim ocsmemo
set ocsmemo = New CCSMemo
	ocsmemo.FCurrPage					= page
	ocsmemo.FPageSize					= 100


if (searchfield = "userid") then
    userid = searchstring
    ocsmemo.FRectUserID = userid
elseif (searchfield = "orderserial") then
    orderserial = searchstring
    ocsmemo.FRectOrderserial = orderserial
elseif (searchfield = "phonenumber") then
    dim phonenumber : phonenumber = searchstring
    ocsmemo.FRectPhoneNumber = phonenumber
''elseif (searchfield = "contents") then
''    dim contents : contents = searchstring
''    ocsmemo.FRectContents = contents
end if

if (contents <> "") then
	ocsmemo.FRectContents = contents
end if

if (finishyn = "N") then
    ocsmemo.FRectIsFinished = "N"
elseif (finishyn = "M") then
    ocsmemo.FRectIsFinished = "N"
    ocsmemo.FRectOrderserial = ""
    ocsmemo.FRectPhoneNumber = ""
    ocsmemo.FRectUserID = ""
    ocsmemo.FRectWriteUser  = session("SSBCtID")
end if

if (finishyn <> "M") then
	ocsmemo.FRectWriteUser = writeuser
end if

ocsmemo.FRectRegStart = LEft(CStr(DateSerial(yyyy1,mm1 ,dd1)),10)
ocsmemo.FRectRegEnd = LEft(CStr(DateSerial(yyyy2,mm2 ,dd2)),10)
ocsmemo.FRectMMGubun = mmgubun
ocsmemo.FRectMMGubunExvlude = MMGubunExvlude
ocsmemo.FRectRetryDateExclude = retrydateexclude
ocsmemo.FRectQadiv = qadiv
ocsmemo.GetCSMemoList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}
function GotoCallHistoryMemoMidify(iid,iorderserial){
    try{
        parent.header.i_ippbxmng.popCallRing('','','',iid,iorderserial,'');
    }catch(e){
        opener.top.header.i_ippbxmng.popCallRing('','','',iid,iorderserial,'');
    }
}

// Not Using
function GotoHistoryMemoMidify(divcd,id,userid,orderserial) {
	var popwin = window.open("/cscenter/history/history_memo_write.asp?divcd=" + divcd + "&id=" + id + "&backwindow=" + "opener" + "&userid=" + userid + "&orderserial=" + orderserial,"GotoHistoryMemoMidify","width=600 height=400 scrollbars=no resizable=no");
	popwin.focus();
}

function OpenOrderMasterDetailWindow(orderserial) {
	var popwin = window.open("/cscenter/ordermaster/ordermaster_detail.asp?orderserial=" + orderserial,"OpenOrderMasterDetailWindow" + orderserial,"width=1300 height=750 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit(frm) {
	/*
	if (frm.contents.value != "") {
		if ((frm.searchfield.value == "") || (frm.searchstring.value == "")) {
			alert("�������� �Է��ؾ߸� ���� �˻��� �����մϴ�.");
			//return;
		}
	}
	*/

	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" height="60" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		������ :
		<select class="select" name="searchfield">
            <option value="" <% if (searchfield = "") then %>selected<% end if %>>����</option>
            <option value="userid" <% if (searchfield = "userid") then %>selected<% end if %>>���̵�</option>
            <option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %>>�ֹ���ȣ</option>
            <option value="phonenumber" <% if (searchfield = "phonenumber") then %>selected<% end if %>>��ȭ��ȣ</option>
			<!--
            <option value="contents" <% if (searchfield = "contents") then %>selected<% end if %>>����</option>
			-->
        </select>
        &nbsp;
        <input type="text" class="text" name="searchstring" value="<%= searchstring %>" size="15" onKeyPress="if (event.keyCode == 13) { jsSubmit(frm); }" >&nbsp;&nbsp;

	    ���� :
	    <input type="text" class="text" name="contents" value="<%= contents %>" size="15" onKeyPress="if (event.keyCode == 13) { jsSubmit(frm); }" >&nbsp;&nbsp;

	    ������ID :
	    <input type="text" class="text" name="writeuser" value="<%= writeuser %>" size="12" onKeyPress="if (event.keyCode == 13) { jsSubmit(frm); }" >&nbsp;&nbsp;
        &nbsp;
        <input type="radio" name="finishyn" value="A" <% if (finishyn = "A") then response.write "checked" end if %>>��ü
        <input type="radio" name="finishyn" value="N" <% if (finishyn = "N") then response.write "checked" end if %>>��ó����û�޸�
	    <input type="radio" name="finishyn" value="M" <% if (finishyn = "M") then response.write "checked" end if %> onClick="frm.searchfield.value='';frm.searchstring.value='';frm.writeuser.value='';"><b>���� ��ó��</b>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
        <!-- #include virtual="/cscenter/memo/mmgubunselectbox.asp"-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		�˻��Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="checkbox" class="checkbox" name="MMGubunExvlude" value="on" <% if (MMGubunExvlude <> "") then %>checked<% end if %>> SMS/�̸��� �ȳ� �޸� ����
		&nbsp;
		<input type="checkbox" class="checkbox" name="retrydateexclude" value="Y" <% if (retrydateexclude <> "") then %>checked<% end if %>> ����ó�� ���� ���� �޸� ����
	</td>
</tr>
</table>
</form>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= FormatNumber(ocsmemo.FTotalCount,0) %></b>
		&nbsp;
		������ : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(ocsmemo.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">idx</td>
    <td width="100">������</td>
    <td width="30">����</td>
    <td width="30">ó��<br>����</td>
	<td width="45">Ư��<br>����</td>
    <td><font color=blue>[���л�]</font><br>����</td>
    <td>������<br>ó����</td>
    <td width="70">������<br>�Ϸ���</td>
    <td width="150">����ó������</td>
    <td width="30">�Ϸ�</td>
	<td width="40">���</td>
</tr>

<% if (ocsmemo.FResultCount > 0) then %>
	<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="#FFFFFF" height="25">
  		<td><%= ocsmemo.FItemList(i).fid %></td>
        <td>
        	<% if ocsmemo.FItemList(i).Fuserid <> "" then %>
        		<%= printUserId(ocsmemo.FItemList(i).Fuserid, 2, "*") %><br>
        	<% end if %>
        	<% if Trim(ocsmemo.FItemList(i).Forderserial) <> "" then %>
        		<a href="javascript:OpenOrderMasterDetailWindow('<%= ocsmemo.FItemList(i).Forderserial %>')"><%= ocsmemo.FItemList(i).Forderserial %></a><br>
        	<% end if %>
        	<%= printtel(ocsmemo.FItemList(i).FphoneNumber) %><br>
        </td>
        <td><%= ocsmemo.FItemList(i).GetmmGubunName %></td>
        <td><%= ocsmemo.FItemList(i).GetDivCDName %></td>
		<td>
			<% if (ocsmemo.FItemList(i).Fspecialmemo = "###") then %><font color="red"><% end if %>
			<%= ocsmemo.FItemList(i).Fspecialmemo %>
		</td>
        <td align="left">
        	<% if ocsmemo.FItemList(i).Fqadiv<>"" then %>
        		<font color="blue">[<%= ocsmemo.FItemList(i).fcomm_name2 %>]</font>
        		<!--<font color="blue">[<%= ocsmemo.FItemList(i).GetQaDivName %>]</font>-->
        	<% else %>
				<font color="blue">[���л󼼾���]</font>
        	<% end if %>

            <a href="javascript:GotoCallHistoryMemoMidify('<%= ocsmemo.FItemList(i).Fid %>','<%= ocsmemo.FItemList(i).Forderserial %>')">
            	<% if Trim(ocsmemo.FItemList(i).Fcontents_jupsu) = "" or isnull(ocsmemo.FItemList(i).Fcontents_jupsu) then %>
            		<br>(�������)
            	<% else %>
            		<br><%= Left(ocsmemo.FItemList(i).Fcontents_jupsu,50) %>
            	<% end if %>
            </a>
        </td>
        <td>
        	<%= ocsmemo.FItemList(i).Fwriteuser %>
        	<% if ocsmemo.FItemList(i).Ffinishyn = "Y" then %>
        		<br><font color=green><%= ocsmemo.FItemList(i).Ffinishuser %></font>
        	<% end if %>
        </td>
        <td align="center">
        	<acronym title="<%= ocsmemo.FItemList(i).Fregdate %>"><%= Left(ocsmemo.FItemList(i).Fregdate,10) %></acronym>
        	<% if ocsmemo.FItemList(i).Ffinishyn = "Y" then %>
        		<br><font color=green><acronym title="<%= ocsmemo.FItemList(i).Ffinishdate %>"><%= Left(ocsmemo.FItemList(i).Ffinishdate,10) %></acronym></font>
        	<% end if %>
        </td>
        <td>
        	<% if (ocsmemo.FItemList(i).Ffinishyn = "N") then %>
        		<%= ocsmemo.FItemList(i).Fretrydate %>
        	<% end if %>
        </td>
        <td>
			<%
			if (ocsmemo.FItemList(i).Ffinishyn = "Y") then
				response.write "�Ϸ�"
			elseif (ocsmemo.FItemList(i).FupchefinishYN = "Y") then
				response.write "<font color='red'>��ü����</font>"
			else
				response.write "<font color='red'>" & ocsmemo.FItemList(i).Ffinishyn & "</font>"
			end if
			%>
        </td>
        <td>
			<input type="button" value="��" onclick="GotoCallHistoryMemoMidify('<%= ocsmemo.FItemList(i).Fid %>','<%= ocsmemo.FItemList(i).Forderserial %>');" class="button">
        </td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="19" align="center" bgcolor="#FFFFFF">
	        <% if ocsmemo.HasPreScroll then %>
			<a href="javascript:goPage('<%= ocsmemo.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + ocsmemo.StartScrollPage to ocsmemo.FScrollCount + ocsmemo.StartScrollPage - 1 %>
	    		<% if i>ocsmemo.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if ocsmemo.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="12">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>

</table>

<script type="text/javascript">

document.onload = getOnload();

function getOnload(){
	startRequest('mmgubun','<%= mmgubun %>','<%= qadiv %>');
}
</script>

<%
set ocsmemo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
