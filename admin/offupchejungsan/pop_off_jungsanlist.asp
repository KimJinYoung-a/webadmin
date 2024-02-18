<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ����
' Hieditor : 2009.04.07 ������ ����
'			 2011.04.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim yyyy1, mm1, dd1 ,taxregdate, isusual, jungsandate, isipkumfinish, dategubun, ipkumdate
dim i, ipsum
	taxregdate      = request("taxregdate")
	isusual         = request("isusual")
	jungsandate     = request("jungsandate")
	isipkumfinish   = request("isipkumfinish")
	dategubun       = request("dategubun")
	yyyy1           = request("yyyy1")
	mm1             = request("mm1")
	dd1             = request("dd1")

if dategubun="" then dategubun="Ipkum" end if
if isipkumfinish="" then isipkumfinish="Y" end if

if taxregdate<>"" then
    yyyy1   = Left(taxregdate,4)
    mm1     = Mid(taxregdate,6,2)
    dd1     = Mid(taxregdate,9,2)
elseif yyyy1<>"" then
    taxregdate = yyyy1 + "-" + mm1 + "-" + dd1
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))

if (taxregdate="") then
    taxregdate = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
end if

if (ipkumdate="") then
    ipkumdate = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
end if

dim ooffjungsan
set ooffjungsan = new COffJungsan

	if dategubun="Ipkum" then
		ooffjungsan.FRectIpkumDate = ipkumdate
	elseif dategubun="Tax" then
	    ooffjungsan.FRectTaxRegDate = taxregdate
	end if
	
	if isusual="Y" then
	    ooffjungsan.FRectGubunCd ="EE"
	elseif isusual="N" then
	    ooffjungsan.FRectGubunCd ="FF"
	end if
	
	if (jungsandate<>"") then
	    ooffjungsan.FRectjungsandate = jungsandate
	end if
	
	if isipkumfinish="Y" then
	    ooffjungsan.FRectfinishflag = "7"
	elseif isipkumfinish="N" then
	    ooffjungsan.FRectfinishflag = "3"
	elseif isipkumfinish="A" then
	    ooffjungsan.FRectfinishflag = "ALL"
	end if
	
	ooffjungsan.JungsanFixedList
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
    	<select name="dategubun" class="select">
    	<option value="Tax" <% if dategubun="Tax" then response.write "selected" %> >������
    	<option value="Ipkum" <% if dategubun="Ipkum" then response.write "selected" %> >�Ա���
    	</select>
    	<% DrawOneDateBox yyyy1, mm1, dd1 %>
    	&nbsp;
    	������౸�� : 
    	<select name="isusual">
    	<option value="" <% if isusual="" then response.write "selected" %> >��ü
    	<option value="Y" <% if isusual="Y" then response.write "selected" %> >�������
    	<option value="N" <% if isusual="N" then response.write "selected" %> >�̿�����
    	</select>
    	&nbsp;
    	
    	������ :
    	<select name="jungsandate">
    	<option value="" <% if jungsandate="" then response.write "selected" %> >��ü
    	<option value="15��" <% if jungsandate="15��" then response.write "selected" %> >15��
    	<option value="����" <% if jungsandate="����" then response.write "selected" %> >����
    	<option value="����" <% if jungsandate="����" then response.write "selected" %> >����
    	</select>
    	&nbsp;
    	�������
    	<select name="isipkumfinish">
    	<option value="A" <% if isipkumfinish="A" then response.write "selected" %> >Ȯ���̻�
    	<option value="N" <% if isipkumfinish="N" then response.write "selected" %> >����Ȯ��
    	<option value="Y" <% if isipkumfinish="Y" then response.write "selected" %> >�ԱݿϷ�
    	</select>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">			
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�����</td>
    <td>������</td>    
    <td>�Ա���</td>
    <td>�귣��ID</td>
    <td>��ü��</td>
    <td>����ڹ�ȣ</td>
    <td>����ݾ�</td>
    <td>�������</td>
    <td>������</td>
    <td>����</td>
    <td>����</td>
    <td>(�������)</td>
    <td>���</td>
</tr>
<% 
if ooffjungsan.FresultCount > 0 then

for i=0 to ooffjungsan.FresultCount-1
%>
<%
ipsum = ipsum + ooffjungsan.FItemList(i).Ftot_jungsanprice
%>

<% if ooffjungsan.FItemList(i).Ftot_jungsanprice<0 then %>
<tr align="center" bgcolor="<%= adminColor("dgray") %>">
<% else %>
<tr align="center" bgcolor="#FFFFFF">
<% end if %>
    <td><%= ooffjungsan.FItemList(i).FYYYYMM %></td>
    <td><%= ooffjungsan.FItemList(i).FTaxRegDate %></td>  
    <td>
        <% if IsNULL(ooffjungsan.FItemList(i).FIpkumdate) or (ooffjungsan.FItemList(i).FIpkumdate="1900-01-01") then %>
        <% else %>
            <%= ooffjungsan.FItemList(i).FIpkumdate %>
        <% end if %>
    </td>
    <td><%= ooffjungsan.FItemList(i).FMakerid %></td>
    <td><%= ooffjungsan.FItemList(i).Fcompany_name %></td>
    <td><%= ooffjungsan.FItemList(i).Fcompany_no %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
    <td><font color="<%= ooffjungsan.FItemList(i).GetStatecolor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
    <td><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
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
    <td>
    (
    <% if ooffjungsan.FItemList(i).Fjungsan_bank = "ȫ�ἧ����" then %>
	HSBC
	<% elseif ooffjungsan.FItemList(i).Fjungsan_bank = "��������" then %>
	����
	<% elseif ooffjungsan.FItemList(i).Fjungsan_bank = "����" then %>
	SC����
	<% elseif ooffjungsan.FItemList(i).Fjungsan_bank = "��Ƽ" then %>
	�ѱ���Ƽ
	<% else %>
	<%= ooffjungsan.FItemList(i).Fjungsan_bank %>
	<% end if %>
    <%= ooffjungsan.FItemList(i).Fjungsan_acctno %>
    )
    </td>
    <td>
  	    <% if Not IsNULL(ooffjungsan.FItemList(i).Fneotaxno) then %>
  	        <img src="/images/icon_print02.gif" width="14" height="14" border=0 onclick="window.open('http://www.bill36524.com/popupBillTax.jsp?NO_TAX=<%= ooffjungsan.FItemList(i).Fneotaxno %>&NO_BIZ_NO=2118700620')" style="cursor:hand">
  	    <% end if %>
	</td>      
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="6"></td>
	<td align="right"><%= FormatNumber(ipsum,0) %></td>
	<td colspan="6"></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan=20>�˻� ����� �����ϴ�</td>
</tr>
<% end if %>
</table>

<%
set ooffjungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->