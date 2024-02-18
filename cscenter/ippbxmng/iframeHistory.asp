<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.10.26 �̻� ����
'           2016.12.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_memocls.asp" -->
<%
dim i, userid, orderserial, phoneNumer, writeuser, finishyn, searchType, searchValue, id
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
	elseif (finishyn = "M") then
	    ocsmemo.FRectIsFinished = "N"
	    ocsmemo.FRectOrderserial = ""
	    ocsmemo.FRectPhoneNumber = ""
	    ocsmemo.FRectUserID = ""
	    ocsmemo.FRectWriteUser  = session("SSBCtID")
	end if

	if (userid <> "") or (orderserial<>"") or (phoneNumer<>"") or (writeuser<>"") or (finishyn="N") or (finishyn="M") then
        ocsmemo.FPageSize = 50
	    ocsmemo.GetCSMemoList
	end if

%>

<script type="text/javascript">

function GotoHistoryMemo(id) {
    <% if InStr(request.ServerVariables("HTTP_REFERER"),"popCallRing.asp")>0 then %>
    parent.location.href = "/cscenter/ippbxmng/popCallRing.asp?id=" + id;
    <% else %>
    parent.location.href = "/cscenter/ippbxmng/CallRingWithOrderFrame.asp?id=" + id;
    <% end if %>
}

function showhideMemo(num, p_totcount)	{
	var i, con;
	for (i = 0; i < p_totcount; i++) {
		con = document.getElementById("Memoblock" + i);

		if (num == i) {
			if (con.style.display == "") {
				con.style.display="none";
			} else {
				con.style.display="";
			}
		}else{
			con.style.display="none";
		}
	}
}

</script>
<link rel="stylesheet" href="/cscenter/css/cs.css" type="text/css">
<style>
body {
    background-color: #FFFFFF;
}

.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}
</style>
<form name="frmSearch" method="get" style="margin:0px;">
<table width="100%" border=0 cellspacing=1 cellpadding=1 class=a bgcolor="EEEEEE">
<tr height="20">
    <td>
        <select name="searchType" class="select">
	        <option value="OD" <%= chkIIF(searchType="OD","selected","") %> >�ֹ���ȣ
	        <option value="UID" <%= chkIIF(searchType="UID","selected","") %> >���̵�
	        <option value="PH" <%= chkIIF(searchType="PH","selected","") %> >��ȭ��ȣ
        </select>
        <input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13">
        <input type="radio" name="finishyn" value="A" <% if (finishyn = "A") then response.write "checked" end if %>>��ü
        <input type="radio" name="finishyn" value="N" <% if (finishyn = "N") then response.write "checked" end if %>>��ó��

        <!-- ���� �ֹ���ȣ/��ȭ��ȣ/���̵� �˻� �ʱ�ȭ �ϸ鼭 ���� �α��� ���̵��� ��ó�� ���� ���� -->
        <input type="radio" name="finishyn" value="M" <% if (finishyn = "M") then response.write "checked" end if %> onClick="frmSearch.searchType.value='';frmSearch.searchValue.value='';"><b>���ǹ�ó��</b>
		<% if C_ADMIN_AUTH then %>
        <br />[������]������:<input type="text" class="text" name="writeuser" value="<%= writeuser %>" size="10">
        <% end if %>
    </td>
    <td width="30" align="right"><input type="button" class="button" value="�˻�" onClick="frmSearch.submit()"></td>
</tr>
</table>
</form>

<table width="100%" border="0" cellpadding="2" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if ocsmemo.FResultCount > 0 then %>
	<tr>
	    <td class="listSep" colspan="6" bgcolor="#CCCCCC" style="border-top:1px"></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="30">����</td>
	 	<td width="80">��ID<br><font color="blue">�ֹ���ȣ</font></td>
		<td><font color="blue">[���л�]</font><br>����</td>
	    <td width="65">�����<br><font color="blue">ó����</font></td>
		<td width="65">������<br><font color="blue">�Ϸ���</font></td>
		<td width="30">�Ϸ�<br>����</td>
	</tr>
	<tr>
	    <td class="listSep" colspan="6" bgcolor="#CCCCCC" style="border-top:1px"></td>
	</tr>

	<% for i = 0 to (ocsmemo.FResultCount - 1) %>
	<tr align="center" bgcolor="<%= chkIIF(CStr(ocsmemo.FItemList(i).Fid)=id,"#DDDDDD","#FFFFFF") %>">
	    <td><%= ocsmemo.FItemList(i).GetDivCDName %><!--<br><%= ocsmemo.FItemList(i).Fid %>--></td>
	 	<td><%= ocsmemo.FItemList(i).Fuserid %><br><font color="blue"><%= ocsmemo.FItemList(i).Forderserial %></font></td>
		<td align="left">
		    <a href="javascript:showhideMemo(<%= i %>,<%= ocsmemo.FResultCount %>);" class="link_ctleft" onFocus="this.blur();">
	        	<% if ocsmemo.FItemList(i).Fqadiv<>"" then %>
	        		<!--<font color="blue">[<%'= getMemoDivName(ocsmemo.FItemList(i).fcomm_name2) %>]</font>-->
	        		<font color="blue">[<%= ocsmemo.FItemList(i).fcomm_name2 %>]</font>
	        	<% else %>
	        		<font color="blue">[���л󼼾���]</font>
	        	<% end if %>

	        	<% if (Replace(Trim(ocsmemo.FItemList(i).Fcontents_jupsu), vbCrLf, "") = "") then %>
	        		<br>(�������)
	        	<% else %>
	        		<br><%= DDotFormat(Replace(ocsmemo.FItemList(i).Fcontents_jupsu, "<", "&lt;"),25) %>.
	        	<% end if %>
	        </a>
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
		<td>
			<%
			if (ocsmemo.FItemList(i).Ffinishyn = "Y") then
				response.write "�Ϸ�"
			elseif (ocsmemo.FItemList(i).FupchefinishYN = "Y") then
				response.write "<font color='red'>��ü����</font>"
			end if
			%>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" id="Memoblock<%= i %>" style="display:none;">
	    <td colspan="6" bgcolor="#F4F9F4"  style="padding:2;border-bottom:1px solid #DCDCDC">
	        <table width="260" border="0" cellspacing="0" cellpadding="0">
	        <tr>
	            <td style="padding:3">
	                <a href="javascript:GotoHistoryMemo('<%= ocsmemo.FItemList(i).Fid %>')">
		            	<% if (Replace(Trim(ocsmemo.FItemList(i).Fcontents_jupsu), vbCrLf, "") = "") then %>
		            		(�������)
		            	<% else %>
		            		<%= Replace(ocsmemo.FItemList(i).Fcontents_jupsu, "<", "&lt;") %>
		            	<% end if %>
	    	        </a>
	            </td>
	        </tr>
	        </table>
	    </td>
	</tr>
	<tr>
	    <td class="listSep" colspan="6" bgcolor="#CCCCCC"></td>
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
