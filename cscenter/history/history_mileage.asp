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
			���Ÿ��ϸ���[<b><%= FormatNumber(CLng(ocsmileage.FItemList(0).Ftotalbuymileage) + CLng(ocsmileage.FItemList(0).Ftotaloldbuymileage),0) %></b>] +
			���ʽ����ϸ���[<b><%= FormatNumber(ocsmileage.FItemList(0).Ftotalbonusmileage,0) %></b>] -
			��븶�ϸ���[<b><%= FormatNumber(ocsmileage.FItemList(0).Ftotalspendmileage,0) %></b>] =
			�ܿ����ϸ���[<b><%= FormatNumber(CLng(ocsmileage.FItemList(0).Ftotalbuymileage) + CLng(ocsmileage.FItemList(0).Ftotaloldbuymileage) + CLng(ocsmileage.FItemList(0).Ftotalbonusmileage) - CLng(ocsmileage.FItemList(0).Ftotalspendmileage),0) %></b>]
			<% else %>
			�˻������ �����ϴ�.(Ż����� �� �ֽ��ϴ�.)
			<% end if %>
        </td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
    <tr height="20" align="center" bgcolor="F3F3FF">
    	<td width="50">IDX</td>
      	<td width="60">���ϸ���</td>
      	<td width="50">����</td>
      	<td width="70">&nbsp;&nbsp;�����ڵ�</td>
      	<td>���䳻��</td>
      	<td width="80">�����</td>
      	<td width="90">�����ֹ���ȣ</td>
      	<td width="60">��������</td>
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
    	    <% if ocsmileagelist.FItemList(i).Fmileage >= 0 then %><font color="blue">����</font><% else %><font color="red">���</font><% end if %>
    	</td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyocd %></td>
    	<td><%= ocsmileagelist.FItemList(i).Fjukyo %></td>
    	<td><acronym title="<%= ocsmileagelist.FItemList(i).Fregdate %>"><%= Left(ocsmileagelist.FItemList(i).Fregdate,10) %></acronym></td>
    	<td><%= ocsmileagelist.FItemList(i).Forderserial %></td>
    	<td><% if (ocsmileagelist.FItemList(i).Fdeleteyn = "Y") then %>����<% end if %></td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>
    <% next %>

<% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td colspan="9">�˻������ �����ϴ�.</td>
    </tr>
<% end if %>
</table>


<%

set ocsmileage = Nothing
set ocsmileagelist = Nothing

%>


<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
