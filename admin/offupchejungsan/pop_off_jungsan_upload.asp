<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
dim research, segumtype
dim thismonth

research = request("research")
segumtype = request("segumtype")


thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)
%>

<script language='javascript'>

function getCSV(searchtype){
    location.href = '/admin/offupchejungsan/pop_off_jungsan_upload_csv.asp?searchtype=' + searchtype;
}

function getExcel(searchtype){
    location.href = '/admin/offupchejungsan/pop_off_jungsan_upload_excel.asp?searchtype=' + searchtype;
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

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectNotIncludeWonChon = "on"
ooffjungsan.FRectYYYYMM = thismonth
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList

dim ipsum,i
ipsum =0
%>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >�ݿ�(<%= thismonth %>) ���ݰ�꼭 (<%= ooffjungsan.FresultCount %>��)</td>
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
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>

	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ooffjungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
        <td><%= Left(ooffjungsan.FItemList(i).Fcompany_name,9) %></td>
         <td><%=ooffjungsan.FItemList(i).Fcompany_no %></td>
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
ooffjungsan.FRectYYYYMM = ""
ooffjungsan.FRectNotIncludeWonChon = "on"
ooffjungsan.FRectNotYYYYMM = thismonth
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList

ipsum =0
%>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >���� ���ݰ�꼭 (<%= ooffjungsan.FresultCount %>��)</td>
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
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>

	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
    	<% if ooffjungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
        <td><%= Left(ooffjungsan.FItemList(i).Fcompany_name,9) %></td>
         <td><%=ooffjungsan.FItemList(i).Fcompany_no %></td>
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
ooffjungsan.FRectYYYYMM = ""
ooffjungsan.FRectNotYYYYMM = ""
ooffjungsan.FRectNotIncludeWonChon = ""
ooffjungsan.FRectOnlyIncludeWonChon = "on"
ooffjungsan.FRectbankingupflag = "Y"

ooffjungsan.JungsanFixedList

ipsum =0
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
    	<td colspan="5" >��õ¡�� ����� (<%= ooffjungsan.FresultCount %>��)</td>
    	<td align=right>
    	  <a href="javascript:getExcel('withholding')"><img src="/images/iexcel.gif" border=0></a>
    	  <a href="javascript:getCSV('withholding')"><img src="/images/icon_arrow_link.gif" border=0></a>
    	</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="60">����</td>
      <td>����</td>
      <td width="80">����ݾ�*0.967</td>
      <td width="120">��ü��</td>
      <td>����ڵ�Ϲ�ȣ</td>
      <td width="120">(��)�ٹ�����</td>
    </tr>
<% for i=0 to ooffjungsan.FresultCount-1 %>
<%
ipsum = ipsum + fix(ooffjungsan.FItemList(i).Ftot_jungsanprice*0.967)
%>

	<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
        <% if ooffjungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then %>
		<td>HSBC</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��������" then %>
		<td>����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "����" then %>
		<td>SC����</td>
		<% elseif ooffjungsan.FItemList(i).Fipkum_bank = "��Ƽ" then %>
		<td>�ѱ���Ƽ</td>
		<% else %>
		<td><%= ooffjungsan.FItemList(i).Fipkum_bank %></td>
		<% end if %>

        <td><%= ooffjungsan.FItemList(i).Fipkum_acctno %></td>
        <td align="right"><%= FormatNumber(fix(ooffjungsan.FItemList(i).Ftot_jungsanprice*0.967),0) %></td>
        <td><%= Left(ooffjungsan.FItemList(i).Fcompany_name,9) %></td>
         <td><%=ooffjungsan.FItemList(i).Fcompany_no %></td>
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
set ooffjungsan = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->