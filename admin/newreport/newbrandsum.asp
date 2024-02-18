<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/newreportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate,toDate
dim searchtype
Dim makerid, ordType, mdid
searchtype = request("searchtype")
if searchtype="" then searchtype="N"

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
makerid = request("makerid")
ordType = request("ordType")
mdid = request("mdid")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))-1
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if


fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CNewReport
oreport.FPageSize = 500
oreport.FRectSearchType = searchtype
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectMakerid = makerid
oreport.FRectOrdType = ordType
oreport.FRectMdid = mdid
oreport.GetNewBrandSellReport

dim i , datelen, datelen2
%>
<script language='javascript'>
function PopUpcheInfo(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="frm" method="get" action="">
<input type="hidden" name="showtype" value="showtype">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td class="a" >
	�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	&nbsp;&nbsp;���MD : <% fnGetMdlist "mdid", mdid %>
	&nbsp;&nbsp;�������� : 
	<select name="ordType" class="select">
		<option value= "">-����-</option>
		<option value= "1" <%= Chkiif(ordType = "1", "selected", "") %> >����ϡ�</option>
		<option value= "2" <%= Chkiif(ordType = "2", "selected", "") %> >����ϡ�</option>
		<option value= "3" <%= Chkiif(ordType = "3", "selected", "") %> >����ո����</option>
		<option value= "4" <%= Chkiif(ordType = "4", "selected", "") %> >����ո����</option>
	</select>
	<br><br>
	�˻��Ⱓ :
	<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	&nbsp;&nbsp;��ü���� :
	<input type=radio name="searchtype" value="N" <% if searchtype="N" then response.write "checked" %> >�űԾ�ü(1�޳� ���)
	<input type=radio name="searchtype" value="A" <% if searchtype="A" then response.write "checked" %> >��ü��ü

	<td class="a" align="right">
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
	</td>
</tr>
</form>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a">
<tr>
	<td>
		* ���ı����� �űԾ�ü-����ո��� ��ü��ü-�Ⱓ�� ����� ����.<br>
		* �˻� ������ �ִ� 500��. ��ǰ���� ���� ���� ����/������� ��ǰ<br>
	</td>
</tr>
</table>

<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td>�귣��ID</td>
	<td>�귣���</td>
	<td>�귣�屸��</td>
	<td>�⺻����</td>
	<td>�����</td>
	<td>���MD</td>
	<td>���<br>����</td>
	<td>��ǰ<br>��</td>
	<td>�����</td>
	<td>���԰�</td>
	<td>����</td>
	<td>����ո���</td>
	<td>�Ⱓ</td>
</tr>
<% for i=0 to oreport.FResultCount - 1 %>
<%
datelen = datediff("d",oreport.FItemList(i).Fregdate,toDate)
datelen2 = datediff("d",fromDate,toDate)

if datelen2<datelen then datelen=datelen2
%>
<tr bgcolor="#FFFFFF">
	<td><a href="javascript:PopUpcheInfo('<%= oreport.FItemList(i).FUserId %>')"><%= oreport.FItemList(i).FUserId %></a></td>
	<td><%= oreport.FItemList(i).Fsocname_kor %></td>
	<td><%= oreport.FItemList(i).GetUserDivName %></td>
	<td><%= oreport.FItemList(i).GetMaeipDivName %> <%= oreport.FItemList(i).Fdefaultmargine %></td>
	<td><%= Left(oreport.FItemList(i).Fregdate,10) %></td>
	<td><%= oreport.FItemList(i).Fmdusername %></td>
	<td><%= oreport.FItemList(i).Fisusing %></td>
	<td align=center><%= oreport.FItemList(i).Fitemcount %></td>
	<td align=right><%= FormatNumber(oreport.FItemList(i).Fsellttl,0) %></td>
	<td align=right><%= FormatNumber(oreport.FItemList(i).Fbuyttl,0) %></td>
	<td align=right><%= FormatNumber(oreport.FItemList(i).Fsellttl-oreport.FItemList(i).Fbuyttl,0) %></td>
	<td align=right>
	<% if datelen<>0 then %>
	<%= FormatNumber(oreport.FItemList(i).Fsellttl/datelen,0) %>
	<% end if %>
	</td>
	<td align=center><%= datelen %></td>
</tr>
<% next %>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->