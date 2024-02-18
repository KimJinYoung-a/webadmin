<%

'==============================================================================
''기 접수된 CS건 있는지 확인
dim oOldcsaslist

set oOldcsaslist = New CCSASList

oOldcsaslist.FRectNotCsID     = id
oOldcsaslist.FRectOrderserial = orderserial

if (orderserial<>"") then
    oOldcsaslist.GetCSASMasterList_3PL
end if

dim ExistsRegedCSCount
ExistsRegedCSCount = oOldcsaslist.FResultCount

%>
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td >
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>CS처리 요청 등록</b></td>
            <td width="140" align="right" <%= ChkIIF(ExistsRegedCSCount>1,"bgcolor='#33CC33'","") %> >
            <% if (ExistsRegedCSCount>1) then %>
                <a href="javascript:ShowOLDCSList();">기 접수된 CS 건 (<%= ExistsRegedCSCount-1 %>)</a>
            <% end if %>
            </td>
        </tr>
        </table>
    </td>
</tr>
<% if (IsDisplayPreviousCSList = true) then %>
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <% for i = 0 to (oOldcsaslist.FResultCount - 1) %>

            <% if CStr(oOldcsaslist.FItemList(i).Fid)<>id then %>
                <% if (oOldcsaslist.FItemList(i).Fdeleteyn = "Y") then %>
                <tr bgcolor="#EEEEEE" style="color:gray" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% else %>
                <tr bgcolor="#FFFFFF" align="center" onclick="ChangeColor(this,'AFEEEE','FFFFFF'); searchDetail('<%= oOldcsaslist.FItemList(i).Fid %>');" style="cursor:hand">
                <% end if %>
                    <td height="20" nowrap><%= oOldcsaslist.FItemList(i).Fid %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).GetAsDivCDName %>"><font color="<%= oOldcsaslist.FItemList(i).GetAsDivCDColor %>"><%= oOldcsaslist.FItemList(i).GetAsDivCDName %></font></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Forderserial %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fmakerid %>"><%= Left(oOldcsaslist.FItemList(i).Fmakerid,32) %></acronym></td>
                    <td nowrap><%= oOldcsaslist.FItemList(i).Fcustomername %></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Fuserid %>"><%= oOldcsaslist.FItemList(i).Fuserid %></acronym></td>
                    <td nowrap align="left"><acronym title="<%= oOldcsaslist.FItemList(i).Ftitle %>"><%= oOldcsaslist.FItemList(i).Ftitle %></acronym></td>
                    <td nowrap><font color="<%= oOldcsaslist.FItemList(i).GetCurrstateColor %>"><%= oOldcsaslist.FItemList(i).GetCurrstateName %></font></td>
                    <td nowrap align="right"><%= FormatNumber(oOldcsaslist.FItemList(i).Frefundrequire,0) %></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Fregdate %>"><%= Left(oOldcsaslist.FItemList(i).Fregdate,10) %></acronym></td>
                    <td nowrap><acronym title="<%= oOldcsaslist.FItemList(i).Ffinishdate %>"><%= Left(oOldcsaslist.FItemList(i).Ffinishdate,10) %></acronym></td>
                    <td nowrap>
                    <% if oOldcsaslist.FItemList(i).Fdeleteyn="Y" then %>
                    <font color="red">삭제</font>
                    <% end if %>
                    </td>
                </tr>
            <% end if %>
        <% next %>
        </table>
    </td>
</tr>
<% end if %>
<%

set oOldcsaslist = Nothing

%>