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
''���� ���ϸ��� �հ�
dim ocsmileage
set ocsmileage = New CCSCenterMileage
ocsmileage.FRectUserID = userid

if (ocsmileage.FRectUserID<>"") then
    ocsmileage.getUserCurrentMileage
end if

'==============================================================================
''���ϸ��� Log
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
''���Ό��  ���ϸ��� �հ�
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���̵� : <input type="text" class="text" name="userid" value="<%= userid %>">
          	&nbsp;
          	<input type="checkbox" name="showall" <%= chkIIF(showall="on","checked","") %> >����������ǥ��
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
          	<input type="button" class="button" value="�˻�" onclick="document.frm.submit()">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>�������</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="85">6��������</td>
    	<td width="85">�ֱ�6����</td>
    	<td width="85">��ī����(����)</td>
    	<td width="85">���Ÿ��ϸ���</td>
    	<td width="85">���ʽ����ϸ���</td>
    	<td width="85">��븶�ϸ���</td>
    	<td width="85">�Ҹ�ȸ��ϸ���</td>
      	<td width="85">�ܿ����ϸ���</td>
      	<td width="110">�Ҹ꿹�� ���ϸ���<br>(<%= oExpireMileTotal.FRectExpireDate %>)</td>
      	<td>���</td>
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


<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <img src="/images/icon_arrow_down.gif" align="absbottom">
		    <strong>���ʽ����ϸ��� �󼼳���</strong>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="120">���̵�</td>
    	<td width="60">Idx</td>
      	<td width="60">���ϸ���</td>
      	<td width="50">����</td>
      	<td width="80">�����ڵ�</td>
      	<td width="200">����</td>
      	<td width="80">�����</td>
      	<td width="90">�ֹ���ȣ</td>
      	<td width="60">��������</td>
      	<td>���</td>
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
    	    <% if ocsmileagelist.FItemList(i).Fmileage >= 0 then %><font color="blue">����</font><% else %><font color="red">���</font><% end if %>
    	</td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyocd %></td>
    	<td align="left"><%= ocsmileagelist.FItemList(i).Fjukyo %></td>
    	<td><acronym title="<%= ocsmileagelist.FItemList(i).Fregdate %>"><%= Left(ocsmileagelist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsmileagelist.FItemList(i).Forderserial %></td>
    	<td><% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>����<% end if %></td>
      	<td></td>
    </tr>
        <% next %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="10"> �˻��� ������ �����ϴ�.</td>
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