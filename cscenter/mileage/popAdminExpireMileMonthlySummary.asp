<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���� ���ϸ����Ҹ�
' History : ������ ����(�⺰ ���ϸ��� �Ҹ�)
'           2023.07.21 �ѿ�� ����(���� ���ϸ��� �Ҹ�� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mileagecls.asp" -->
<%
dim userid, yyyymmdd, myMileage, oExpireMile, oExpireMileTotal, currentdate, i, menupos
dim Tot_GainMileage, Tot_MonthMaySpendMileage, Tot_MonthMayRemainMileage, Tot_realExpiredMileage
    userid = requestcheckvar(request("userid"),32)
    yyyymmdd = requestCheckvar(request("yyyymmdd"),10)
    menupos = requestCheckvar(getNumeric(request("menupos")),10)

currentdate=date()

if (yyyymmdd="") then
    ' �̹��޸���
    yyyymmdd=dateadd("d",-1,dateserial(year(dateadd("m",+1,currentdate)),month(dateadd("m",+1,currentdate)),"01"))
end if

' ���� ���ϸ���
set myMileage = new CCSCenterMileage
    myMileage.FRectUserID = userid
    if (userid<>"") then
        myMileage.getUserCurrentMileage
    end if

' ���Ό�� ���ϸ��� �⵵�� ����Ʈ
set oExpireMile = new CCSCenterMileage
    oExpireMile.FRectUserid = userid
    ' �ش�Expire ������ ������ ���
    ' oExpireMile.FRectExpireDate = yyyymmdd

    if (userid<>"") then
        oExpireMile.getNextExpireMileageMonthlyList
    end if

''���Ό��  ���ϸ��� �հ�
set oExpireMileTotal = new CCSCenterMileage
    oExpireMileTotal.FRectUserid = userid
    oExpireMileTotal.FRectExpireDate = yyyymmdd

    if (userid<>"") then
        oExpireMileTotal.getNextExpireMileageMonthlySum
    end if

%>
<style>
.black12px {font-family: ����; FONT-SIZE: 12px; COLOR: #000000;  TEXT-DECORATION: none; font-weight: bold;}
</style>
<script type='text/javascript'>

function research(frm){
    if (frm.userid.value.length<1){
        alert('���̵� �Է��ϼ���.');
        frm.userid.focus();
        return;
    }
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frmresearch" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="yyyymmdd" value="<%= yyyymmdd %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
    <td align="left">
        ���̵� : <input type="text" name="userid" value="<%= userid %>" size="16" maxlength="32" class="text">
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="�˻�" onClick="research(frmresearch);">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left">

    </td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        �� ���ϸ����� �ο��� ������ ���Ǹ� ���� ��, 60���� �� �̻�� �� 60������ �Ǵ� �� ���Ͽ� �ڵ� �Ҹ�˴ϴ�.
        <br>��) <%= left(dateadd("m",-60,DateSerial(Year(yyyymmdd), month(yyyymmdd), day(yyyymmdd))),7) %>�� 
        ���� ���ϸ��� 4500 / ��� ���ϸ��� 4000 / �ܿ� ���ϸ��� 500 �� ��� 500 ����Ʈ�� <%= yyyymmdd %>�� �ڵ� �Ҹ�˴ϴ�.
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>������¥</td>
    <td>�������ϸ���</td>
    <td>���</td>
    <td>�Ҹ�</td>
    <td>�ܿ�</td>
    <td>�Ҹ꿹����</td>
</tr>
<% if oExpireMile.FResultCount>0 then %>
<%
for i=0 to oExpireMile.FResultCount-1

Tot_GainMileage           = Tot_GainMileage + oExpireMile.FItemList(i).getGainMileage
Tot_MonthMaySpendMileage   = Tot_MonthMaySpendMileage + oExpireMile.FItemList(i).getMonthlyMaySpendMileage
Tot_MonthMayRemainMileage  = Tot_MonthMayRemainMileage + oExpireMile.FItemList(i).getMonthlyMayRemainMileage
Tot_realExpiredMileage    = Tot_realExpiredMileage + oExpireMile.FItemList(i).FrealExpiredMileage
%>
<tr align="center" bgcolor="#FFFFFF">
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%
        ' 2018��8�� ���ϸ��� ���� ������ ���ϸ��� �Ҹ��� ������. ���� �����ʹ� ����� ��������.
        if datediff("d",oExpireMile.FItemList(i).Fregmonth&"-01","2018-08-01")>0 then
        %>
            <%= left(oExpireMile.FItemList(i).Fregmonth,4) %>
        <% else %>
            <%= oExpireMile.FItemList(i).Fregmonth %>
        <% end if %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getGainMileage ,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getMonthlyMaySpendMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).FrealExpiredMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= FormatNumber(oExpireMile.FItemList(i).getMonthlyMayRemainMileage,0) %>
    </td>
    <td <%= chkIIF(yyyymmdd=CStr(oExpireMile.FItemList(i).FExpiredate),"class='black12px'","") %> >
        <%= oExpireMile.FItemList(i).FExpiredate %>
    </td>
</tr>   
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="16" align="center" class="page_link">[ <%= yyyymmdd %> �Ҹ� ��� ������ �����ϴ�.]</td>
    </tr>
<% end if %>

<%
' ���� ���ϸ������� ������ ���.
if (oExpireMile.FResultCount>0) and (oExpireMile.FRectExpireDate="") then
%>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td><%= left(dateadd("m",-59,DateSerial(Year(yyyymmdd), month(yyyymmdd), day(yyyymmdd))),7) & "��~" %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusMileage - Tot_GainMileage ,0) %>
        <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage - Tot_MonthMaySpendMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FrealExpiredMileage - Tot_realExpiredMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage - Tot_MonthMayRemainMileage,0) %></td>
        <td></td>
    </tr>
    <tr height="1" bgcolor="#FFFFFF">
        <td colspan="6"></td>
    </tr>
    <tr bgcolor="#FFFFFF" align="center" height="26">
        <td>�հ�</td>
        <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusMileage,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage ,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.FrealExpiredMileage ,0) %></td>
        <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage,0) %></td>
        <td>&nbsp;</td>
    </tr>
    <% if myMileage.FResultCount>0 then %>
        <tr height="1" bgcolor="#FFFFFF">
            <td colspan="6"></td>
        </tr>
        <tr bgcolor="#FFFFFF" align="center" height="26">
            <td>����</td>
            <td><%= FormatNumber(myMileage.FOneItem.FJumunMileage + myMileage.FOneItem.FFlowerJumunmileage + myMileage.FOneItem.FAcademymileage + myMileage.FOneItem.FBonusmileage,0) %></td>
            <td><%= FormatNumber(myMileage.FOneItem.FSpendMileage ,0) %></td>
            <td>&nbsp;</td>
            <td><%= FormatNumber(myMileage.FOneItem.getCurrentMileage,0) %></td>
            <td>&nbsp;</td>
        </tr>
    <% end if %>
<% end if %>
</table>

<br>
<table width="600" border="0" cellpadding="2" cellspacing="1" bgcolor="d3d3d3" class="a">
<tr bgcolor="#FFFFFF" align="center" height="26">
    <td align="center" >
    <font style="font-family: ����; COLOR: #333333; FONT-SIZE: 13px; font-weight: bold;">
    <%= oExpireMileTotal.FOneItem.getKorExpireDateStr %> �Ҹ� ��� ���ϸ��� : <font color="red"><%= FormatNumber(oExpireMileTotal.FOneItem.getMayExpireTotal,0) %> </font> Point
    </font>
    </td>
</tr>
</table>

<%
set myMileage = Nothing
set oExpireMile = Nothing
set oExpireMileTotal = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->
