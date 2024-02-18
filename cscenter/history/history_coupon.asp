<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->


<%

dim i, userid

userid = request("userid")

'==============================================================================
dim ocscoupon
set ocscoupon = New CCSCenterCoupon

ocscoupon.FRectUserID = userid

if (userid<>"") then
    ocscoupon.GetCSCenterCouponList
end if
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
<table width="100%" border=0 cellspacing=0 cellpadding=2 class=a bgcolor="FFFFFF" align="center">
<% if (userid<>"") then %>
    <tr align="center" bgcolor="#F3F3FF">
        <td height="20" width="30">idx</td>
        <td>������</td>
        <td width="60">���ΰ�</td>
        <td width="80">�ּұ��űݾ�</td>
        <td width="140">��ȿ�Ⱓ</td>
        <td width="65">�����</td>
        <td width="30">���</td>
        <td width="75">����ֹ���ȣ</td>
        <td width="30">����</td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC" style="border-top:1px"></td>
    </tr>
    <% for i = 0 to (ocscoupon.FResultCount - 1) %>
    <tr align="center" <% if (ocscoupon.FItemList(i).Fdeleteyn = "Y") then %>bgcolor="#EEEEEE" class="gray"<% else %>bgcolor="#FFFFFF"<% end if %>>
        <td height="20"><%= ocscoupon.FItemList(i).Fmasteridx %></td>
        <td align="left"><%= ocscoupon.FItemList(i).Fcouponname %></td>
        <td><%= FormatNumber(ocscoupon.FItemList(i).Fcouponvalue,0) %><%= ocscoupon.FItemList(i).GetCouponTypeUnit %></td>
        <td><%= FormatNumber(ocscoupon.FItemList(i).Fminbuyprice,0) %></td>
        <td><acronym title="<%= ocscoupon.FItemList(i).Fstartdate %>"><%= Left(ocscoupon.FItemList(i).Fstartdate,10) %></acronym> ~ <acronym title="<%= ocscoupon.FItemList(i).Fexpiredate %>"><%= Left(ocscoupon.FItemList(i).Fexpiredate,10) %></acronym></td>
        <td><acronym title="<%= ocscoupon.FItemList(i).Fregdate %>"><%= Left(ocscoupon.FItemList(i).Fregdate,10) %></acronym></td>
        <td><% if (ocscoupon.FItemList(i).Fisusing = "Y") then %>���<% end if %></td>
        <td><%= ocscoupon.FItemList(i).Forderserial %></td>
        <td><% if (ocscoupon.FItemList(i).Fdeleteyn = "Y") then %>����<% end if %></td>
    </tr>
    <tr>
        <td class="listSep" colspan="15" bgcolor="#CCCCCC"></td>
    </tr>
    <% next %>
    <% if (ocscoupon.FResultCount < 1) then %>
    <tr bgcolor="#FFFFFF" align="center">
        <td colspan="14">�˻������ �����ϴ�.</td>
    </tr>
    <% end if %>
<% else %>
    <tr align="center" bgcolor="#FFFFFF">
        <td colspan="14" align="center">[��ȿ�� UserID�� �ƴմϴ�.]</td>
    </tr>
<% end if %>
</table>

<%

set ocscoupon = Nothing

%>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
