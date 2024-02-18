<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���庸��
' History : 2015.05.27 ������ ����
'			2023.04.26 �ѿ�� ����(����¡�� �ӽ� ����)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/etcsongjangcls.asp"-->
<%
dim idarr
idarr = request("idarr")
idarr = Mid(idarr,2,Len(idarr))
idarr = replace(idarr,"|",",")

dim osongjang

set osongjang = new CEventsBeasong
osongjang.getEventSongJangList idarr

dim i, bufstr
%>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	�Ѱ˻��Ǽ� : <%= (osongjang.FTotalcount) %>
        </td>
        <td align="right">

        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="70">������ȣ</td>
    	<td width="80">���̵�</td>
    	<td width="50">����</td>
    	<td width="50">������</td>
    	<td width="80">��ȭ��ȣ</td>
    	<td width="80">�ڵ�����ȣ</td>
    	<td width="60">�����ȣ</td>
    	<td width="100">�ּ�1</td>
    	<td width="100">�ּ�2</td>
      	<td width="100">�̺�Ʈ��</td>
      	<td width="100">��ǰ��</td>
      	<td width="30">����</td>
      	<td>��Ÿ����</td>
    </tr>
<% if (osongjang.FTotalcount)<1 then %>
    <tr align="center" bgcolor="#FFFFFF">
  		<td colspan="21" align="center">�˻������ �����ϴ�.</td>
    </tr>
<% else %>
    <% for i=0 to osongjang.FTotalcount -1 %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osongjang.FItemList(i).Fsongjangno %></td>
    	<td><%= osongjang.FItemList(i).Fuserid %></td>
    	<td><%= osongjang.FItemList(i).FuserName %></td>
    	<td><%= osongjang.FItemList(i).FreqName %></td>
    	<td><%= osongjang.FItemList(i).Freqphone %></td>
    	<td><%= osongjang.FItemList(i).Freqhp %></td>
    	<td><%= osongjang.FItemList(i).Freqzipcode %></td>
    	<td><%= osongjang.FItemList(i).Freqaddress1 %></td>
    	<td><%= osongjang.FItemList(i).Freqaddress2 %></td>
      	<td><%= osongjang.FItemList(i).Fgubunname %></td>
      	<td><%= osongjang.FItemList(i).getPrizeTitle  %></td>
      	<td></td>
      	<td><%= osongjang.FItemList(i).Freqetc %></td>
    </tr>
	<%
	if i mod 300 = 0 then
		Response.Flush		' ���۸��÷���
	end if

	next
	%>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
    

<%
set osongjang = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->