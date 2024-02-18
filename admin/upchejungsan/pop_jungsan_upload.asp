<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim research, segumtype
dim thismonth

research = request("research")
segumtype = request("segumtype")


thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)
%>

<script language='javascript'>

function getCSV(searchtype){
    location.href = '/admin/upchejungsan/pop_jungsan_upload_csv.asp?searchtype=' + searchtype;
}

function getExcel(searchtype){
    location.href = '/admin/upchejungsan/pop_jungsan_upload_excel.asp?searchtype=' + searchtype;
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<img src="/images/icon_star.gif" align="absbottom"> <strong>���곻�� ���ε帮��Ʈ</strong>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<%

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

dim ipsum,i
ipsum =0
%>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >�ݿ�(<%= thismonth %>) ���ݰ�꼭 (<%= ojungsan.FresultCount %>��)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('thismonth')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('thismonth')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="60">����</td>
        <td>����</td>
        <td width="80">����ݾ�</td>
        <td width="120">��ü��</td>
        <td>����ڵ�Ϲ�ȣ</td>
        <td width="120">(��)�ٹ�����</td>
    </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ojungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(��)�ٹ�����</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = "on"
ojungsan.FRectNotYYYYMM = thismonth
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

ipsum =0
%>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >���� ���ݰ�꼭 (<%= ojungsan.FresultCount %>��)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('prevmonth')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('prevmonth')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="60">����</td>
      <td>����</td>
      <td width="80">����ݾ�</td>
      <td width="120">��ü��</td>
      <td>����ڵ�Ϲ�ȣ</td>
      <td width="120">(��)�ٹ�����</td>
     </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalSuplycash
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
    	<% if ojungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(��)�ٹ�����</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
ojungsan.FRectYYYYMM = ""
ojungsan.FRectNotYYYYMM = ""
ojungsan.FRectNotIncludeWonChon = ""
ojungsan.FRectOnlyIncludeWonChon = "on"
ojungsan.FRectbankingupflag = "Y"

ojungsan.JungsanFixedList

ipsum =0
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >��õ¡�� ����� (<%= ojungsan.FresultCount %>��)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('withholding')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('withholding')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="60">����</td>
      <td>����</td>
      <td width="80">����ݾ�</td>
      <td width="120">��ü��</td>
      <td>����ڵ�Ϲ�ȣ</td>
      <td width="120">(��)�ٹ�����</td>
    </tr>
<% for i=0 to ojungsan.FresultCount-1 %>
<%
ipsum = ipsum + ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum
%>

	<% if ojungsan.FItemList(i).GetTotalSuplycash<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ojungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ojungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ojungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ojungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalWithHoldingJungSanSum,0) %></td>
        <td><%= Left(ojungsan.FItemList(i).Fcompany_name,9) %></td>
        <td><%=ojungsan.FItemList(i).Fcompany_no%></td>
        <td>(��)�ٹ�����</td>
	</tr>
<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"></td>
		<td align="right"><%= FormatNumber(ipsum,0) %></td>
		<td colspan="3"></td>
	</tr>
</table>

<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->