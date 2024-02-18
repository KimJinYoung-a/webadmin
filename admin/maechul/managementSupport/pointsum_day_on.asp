<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� ����Ʈ ���
' History : 2013.01.14 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/pointsum_on_cls.asp" -->

<%
Dim i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate, jukyocd
dim cgainlog, cspendlog, cofflineshift, cuseroutpoint, cdelpoint
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	jukyocd     = request("jukyocd")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = Cstr(day( dateadd("m",-1,date()) ))	
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2 +1)

Set cgainlog = New cpointsum_on_list
	cgainlog.FRectStartdate = fromDate
	cgainlog.FRectEndDate = toDate
	cgainlog.FPageSize = 100
	cgainlog.FCurrPage	= 1
	
	'//������
	if jukyocd="gainpoint" and jukyocd<>"" then
		cgainlog.fpointsum_gainlog_list_on()
	end if

Set cspendlog = New cpointsum_on_list
	cspendlog.FRectStartdate = fromDate
	cspendlog.FRectEndDate = toDate
	cspendlog.FPageSize = 100
	cspendlog.FCurrPage	= 1
	
	'//������
	if jukyocd="spendpoint" and jukyocd<>"" then
		cspendlog.fpointsum_spendlog_list_on()
	end if

Set cofflineshift = New cpointsum_on_list
	cofflineshift.FRectStartdate = fromDate
	cofflineshift.FRectEndDate = toDate
	cofflineshift.FPageSize = 100
	cofflineshift.FCurrPage	= 1
	
	'//����������ȯ
	if jukyocd="offlineshiftpoint" and jukyocd<>"" then
		cofflineshift.fpointsum_offlineshiftlog_list_on()
	end if

Set cuseroutpoint = New cpointsum_on_list
	cuseroutpoint.FRectStartdate = fromDate
	cuseroutpoint.FRectEndDate = toDate
	cuseroutpoint.FPageSize = 100
	cuseroutpoint.FCurrPage	= 1
	
	'//ȸ��Ż��
	if jukyocd="useroutpoint" and jukyocd<>"" then
		cuseroutpoint.fpointsum_useroutpointlog_list_on()
	end if

Set cdelpoint = New cpointsum_on_list
	cdelpoint.FRectStartdate = fromDate
	cdelpoint.FRectEndDate = toDate
	cdelpoint.FPageSize = 100
	cdelpoint.FCurrPage	= 1
	
	'//�Ҹ�
	if jukyocd="delpoint" and jukyocd<>"" then
		cdelpoint.fpointsum_delpoint_list_on()
	end if
		

if jukyocd="" then
	response.write "<script language='javascript'>"
	response.write "	alert('����Ʈ������ �������ּ���');"
	response.write "</script>"
end if
%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				�Ⱓ : <% DrawDateBoxdynamic yyyy1,"yyyy1",yyyy2,"yyyy2",mm1,"mm1",mm2,"mm2",dd1,"dd1",dd2,"dd2" %>
				����Ʈ���� : <% drawjukyocd_on "jukyocd",jukyocd," onchange='searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>	
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<font color="red">�� ������ ���ϰ� ū ������ �Դϴ�. �Ѵ޴��� �̻� �˻��� ������ �ּ���.</font>
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
dim onpoint, onpointAca
	onpoint = 0
	onpointAca = 0
%>
<% if cgainlog.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		������ �˻���� : <b><%= cgainlog.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>������(10x10)</td>
    <td>������(ACA)</td>
</tr>
<%
For i = 0 To cgainlog.FresultCount -1

onpoint = onpoint + clng(cgainlog.fitemlist(i).fgainpoint)
onpointAca = onpointAca + clng(cgainlog.fitemlist(i).FacademyGainPoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cgainlog.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cgainlog.fitemlist(i).fgainpoint,0) %>
	</td>
	<td>
		<%= FormatNumber(cgainlog.fitemlist(i).FacademyGainPoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cspendlog.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		������ �˻���� : <b><%= cspendlog.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>������(10x10)</td>
    <td>������(ACA)</td>
</tr>
<%
For i = 0 To cspendlog.FresultCount -1

onpoint = onpoint + clng(cspendlog.fitemlist(i).fspendpoint)
onpointAca = onpointAca + clng(cspendlog.fitemlist(i).FacademySpendPoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cspendlog.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cspendlog.fitemlist(i).fspendpoint,0) %>
	</td>
	<td>
		<%= FormatNumber(cspendlog.fitemlist(i).FacademySpendPoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cofflineshift.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		����������ȯ �˻���� : <b><%= cofflineshift.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>����������ȯ</td>
</tr>
<%
For i = 0 To cofflineshift.FresultCount -1

onpoint = onpoint + clng(cofflineshift.fitemlist(i).fofflineshift)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cofflineshift.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cofflineshift.fitemlist(i).fofflineshift,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cuseroutpoint.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		ȸ��Ż�� �˻���� : <b><%= cuseroutpoint.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>ȸ��Ż��</td>
</tr>
<%
For i = 0 To cuseroutpoint.FresultCount -1

onpoint = onpoint + clng(cuseroutpoint.fitemlist(i).fuseroutpoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cuseroutpoint.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cuseroutpoint.fitemlist(i).fuseroutpoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if cdelpoint.FresultCount > 0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�Ҹ� �˻���� : <b><%= cdelpoint.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>�Ҹ�</td>
</tr>
<%
For i = 0 To cdelpoint.FresultCount -1

onpoint = onpoint + clng(cdelpoint.fitemlist(i).fdelpoint)
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= left(cdelpoint.fitemlist(i).fyyyymmdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cdelpoint.fitemlist(i).fdelpoint,0) %>
	</td>
</tr>	
<% next %>
<% end if %>

<% if onpoint <> 0 then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>
		 �հ�
	</td>
	<td>
		<%= FormatNumber(onpoint,0) %>
	</td>
	<% if jukyocd="gainpoint" or jukyocd="spendpoint" then %>
	<td>
	<%= FormatNumber(onpointACA,0) %>
	</td>
	<% end if %>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<% 
Set cgainlog = Nothing
set cspendlog = nothing
set cofflineshift = nothing
set cuseroutpoint = nothing
set cdelpoint = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->