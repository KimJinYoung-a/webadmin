<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs �޸�
' History : 2007.10.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenterv2/lib/classes/history/cs_memocls.asp" -->
<%

dim i, userid, orderserial, phoneNumer, writeuser, finishyn
dim searchType, searchValue
dim id, sitename
userid      = requestCheckVar(request("userid"),32)
orderserial = requestCheckVar(request("orderserial"),32)
phoneNumer  = requestCheckVar(request("phoneNumer"),32)

searchType = requestCheckVar(request("searchType"),32)
searchValue = requestCheckVar(request("searchValue"),32)

writeuser = requestCheckVar(request("writeuser"),32)
finishyn  = requestCheckVar(request("finishyn"),32)
id        = requestCheckVar(request("id"),32)
sitename  = requestCheckVar(request("sitename"),32)

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



'==============================================================================
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
	ocsmemo.FRectSiteName = sitename
    ocsmemo.GetCSMemoList
end if

%>
<script language='javascript'>
function GotoHistoryMemo(id) {
    <% if InStr(request.ServerVariables("HTTP_REFERER"),"popCallRing.asp")>0 then %>
    parent.location.href = "/cscenterv2/order/popCallRing.asp?id=" + id;
    <% else %>
    parent.location.href = "/cscenterv2/order/CallRingWithOrderFrame.asp?id=" + id;
    <% end if %>
}

function showhideMemo(num, p_totcount)	{
  for (i=0; i<p_totcount; i++)   {
	  menu=eval("document.all.Memoblock"+i+".style");

	  if (num==i ){
		if (menu.display=="block"){
			menu.display="none";
		}else{
		  menu.display="block";
		}
	  }else{
		 menu.display="none";
	  }
	}
}

</script>
<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border=0 cellspacing=1 cellpadding=1 class=a bgcolor="EEEEEE">
	<form name="frmSearch" method="get">
	<input type="hidden" name="sitename" value="<%= sitename %>">
	<tr height="20">
	    <td>
	        <select name="searchType" class="select">
		        <option value="OD" <%= chkIIF(searchType="OD","selected","") %> >�ֹ���ȣ
		        <option value="UID" <%= chkIIF(searchType="UID","selected","") %> >���̵�
		        <option value="PH" <%= chkIIF(searchType="PH","selected","") %> >��ȭ��ȣ
	        </select>
	        <input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="13">
	        <!--
	        ������:<input type="text" class="text" name="writeuser" value="<%= writeuser %>" size="10">
	        -->
            <input type="radio" name="finishyn" value="A" <% if (finishyn = "A") then response.write "checked" end if %>>��ü
            <input type="radio" name="finishyn" value="N" <% if (finishyn = "N") then response.write "checked" end if %>>��ó��

            <!-- ���� �ֹ���ȣ/��ȭ��ȣ/���̵� �˻� �ʱ�ȭ �ϸ鼭 ���� �α��� ���̵��� ��ó�� ���� ���� -->
            <input type="radio" name="finishyn" value="M" <% if (finishyn = "M") then response.write "checked" end if %> onClick="frmSearch.searchType.value='';frmSearch.searchValue.value='';"><b>���ǹ�ó��</b>
	    </td>
	    <td width="30" align="right"><input type="button" class="button" value="�˻�" onClick="frmSearch.submit()"></td>
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
    	<td><font color="blue">[���л�]</font><br>����</td>
        <td width="65">�����<br><font color="blue">ó����</font></td>
    	<td width="65">������<br><font color="blue">�Ϸ���</font></td>
    	<td width="30">�Ϸ�<br>����</td>
    </tr>
    <tr>
        <td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
    </tr>

<% for i = 0 to (ocsmemo.FResultCount - 1) %>
    <tr align="center" bgcolor="<%= chkIIF(CStr(ocsmemo.FItemList(i).Fid)=id,"#DDDDDD","#FFFFFF") %>">
        <td><%= ocsmemo.FItemList(i).GetDivCDName %><!--<br><%= ocsmemo.FItemList(i).Fid %>--></td>
     	<td><%= ocsmemo.FItemList(i).Fuserid %><br><font color="blue"><%= ocsmemo.FItemList(i).Forderserial %></font></td>
    	<td align="left">
    	    <a href="javascript:showhideMemo(<%= i %>,<%= ocsmemo.FResultCount %>);" class="link_ctleft" onFocus="this.blur();">
            	<font color="blue">[<%= getMemoDivName(ocsmemo.FItemList(i).Fqadiv) %>]</font><br>
            	<% if (Replace(Trim(ocsmemo.FItemList(i).Fcontents_jupsu), vbCrLf, "") = "") then %>
            		(�������)
            	<% else %>
            		<%= DDotFormat(Replace(ocsmemo.FItemList(i).Fcontents_jupsu, "<", "&lt;"),25) %>.
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
    	<td><% if (ocsmemo.FItemList(i).Ffinishyn = "Y") then %>�Ϸ�<% end if %></td>
    </tr>
    <tr bgcolor="#FFFFFF" id="Memoblock<%= i %>" style="DISPLAY:none;">
        <td colspan="2"></td>
        <td colspan="5" bgcolor="#F4F9F4"  style="padding:2;border-bottom:1px solid #DCDCDC">
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
