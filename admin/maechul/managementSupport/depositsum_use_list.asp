<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��ġ�� ���� ���
' History : 2012.12.05 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/deposit/depositsum_cls.asp" -->

<%
Dim i, yyyy1, mm1, dd1, yyyy2, mm2, dd2, fromDate, toDate, cuse, jukyocd, onoffgubun
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	onoffgubun     = request("onoffgubun")	
	jukyocd     = request("jukyocd")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd1="") then dd1 = Cstr(day( dateadd("m",-1,date()) ))	
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
	
fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2 +1)

if onoffgubun = "" then onoffgubun = "ONLINE"
	
Set cuse = New cdepositsum_list
	cuse.FRectStartdate = fromDate
	cuse.FRectEndDate = toDate
	cuse.FPageSize = 1000
	cuse.FCurrPage	= 1
	cuse.frectjukyocd = jukyocd
	cuse.FRectonoffgubun = onoffgubun	
	cuse.fdepositsum_use_list()

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
				��뱸�� : <% drawonoffgubun "onoffgubun",onoffgubun," onchange='javascript:searchSubmit();'" %>
				���Ű������� : <% drawdepositjukyocd "jukyocd",jukyocd," onchange='searchSubmit();'" %>
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
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= cuse.FresultCount %></b> �� �� 1000�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>�ֹ���ȣ</td>
    <td>��¥</td>
    <td>�ݾ�</td>    
    <td>����ó</td>    
    <td>�����<Br>���̵�</td>
    <td>����<Br>��������</td>
    <td>��������</td>        
</tr>
<%
dim useCash
	useCash = 0
	
if cuse.FresultCount > 0 then
	
For i = 0 To cuse.FresultCount -1

useCash = useCash + cuse.fitemlist(i).fdeposit
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= cuse.fitemlist(i).forderserial %>
	</td>
	<td>
		<%= left(cuse.fitemlist(i).fYYYYMMdd,10) %>
	</td>
	<td>
		<%= FormatNumber(cuse.fitemlist(i).fdeposit,0) %>
	</td>
	<td>
		<%= cuse.fitemlist(i).fsitename %><%= cuse.fitemlist(i).fshopid %>
	</td>	
	<td>
		<%= cuse.fitemlist(i).fuserid %>
	</td>	
	<td>
		<%= cuse.fitemlist(i).fjukyo %>
	</td>
	<td>
		<%= cuse.fitemlist(i).fdeleteYn %>
	</td>
</tr>	
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan=2>
		 �հ�
	</td>
	<td>
		<%= FormatNumber(useCash,0) %>
	</td>
		
	<td colspan=5></td>
</tr>	

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<% 
Set cuse = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->