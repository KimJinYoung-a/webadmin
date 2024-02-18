<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/ipkumlistcls.asp"-->
<%

dim yyyy1,mm1,dd1
dim yyyy2,mm2,dd2
dim ipkumstate,tenbank,ipkumname,page

ipkumstate=request("ipkumstate")
tenbank=request("tenbank")
ipkumname=request("ipkumname")
page=request("page")
if page="" then page=1

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if 

dim ipkum,i,ix
set ipkum = new IpkumChecklist

ipkum.FCurrpage=page
ipkum.FPagesize=150
ipkum.FScrollCount = 5
ipkum.ipkumstate=ipkumstate
ipkum.Ctenbank=tenbank
ipkum.ipkumname=ipkumname

ipkum.yyyy1=yyyy1
ipkum.mm1=mm1
ipkum.dd1=dd1
ipkum.yyyy2=yyyy2
ipkum.mm2=mm2
ipkum.dd2=dd2

ipkum.Getipkumlist


%>
<script language='javascript'>

function ExcelSheet(){
	var b=document.frmipkum.tenbank.value;
	var n=document.frmipkum.ipkumname.value
	var yy1=document.frmipkum.yyyy1.value;
	var mm1=document.frmipkum.mm1.value;
	var dd1=document.frmipkum.dd1.value;
	var yy2=document.frmipkum.yyyy2.value;
	var mm2=document.frmipkum.mm2.value;
	var dd2=document.frmipkum.dd2.value;
	
	window.open('popipkumsheet.asp?yyyy1=' + yy1 + '&yyyy2=' +  yy2 + '&mm1=' + mm1 + '&mm2=' + mm2  + '&dd1=' + dd1 + '&dd2=' + dd2 + '&tenbank=' +  b + '&ipkumname=' + n + '&xl=on');
}
 function scrollmove(v) {
 	document.frmipkum.page.value=v;
 	document.frmipkum.action='ipkumlist.asp';
 	document.frmipkum.submit();
 	
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frmipkum" method="get" action="">
		<input type="hidden" name="showtype" value="showtype">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="page" value="">
	<tr>
		<td class="a" >
		Ȯ�α��� :
		<select name="ipkumstate">
		<option value="">��ü
		<option value="1" <% if ipkumstate="1" then response.write " selected" %>>�Աݹ�Ȯ��
		<option value="0" <% if ipkumstate="0" then response.write " selected" %>>��ó��
		</select>
		���� :
		<select name="tenbank">
		<option value="">��ü
		<option value="����" <% if tenbank="����" then response.write " selected" %>>����
		<option value="����" <% if tenbank="����" then response.write " selected" %>>����
		<option value="�츮" <% if tenbank="�츮" then response.write " selected" %>>�츮
		<option value="�ϳ�" <% if tenbank="�ϳ�" then response.write " selected" %>>�ϳ�
		<option value="���" <% if tenbank="���" then response.write " selected" %>>���
		<option value="����" <% if tenbank="����" then response.write " selected" %>>����
		</select>
		�˻��Ⱓ :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<br>
		�Ա��ڸ� :
		<input type=text name=ipkumname value="<%= ipkumname %>" size=10 >
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
	</form>
</table>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="a">
<tr>
	<td align=right><a href="javascript:ExcelSheet();"><img src="/images/iexcel.gif" border=0 align=absmiddle>��������</a></td>
</tr>
</table>
<table width="800" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td align=center>Idx</td>
	<td align=center>����</td>
	<td align=center>��¥</td>
	<td align=center>����</td>
	<td align=center>�Ա���</td>
	<td align=center>��ݾ�</td>
	<td align=center>�Աݾ�</td>
	<td align=center>�ܾ�</td>
	<td align=center>����</td>
</tr>
<% if ipkum.FResultCount<1 then %>
<% else %>
<% for i=0 to ipkum.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td align=center><%= ipkum.Fipkumitem(i).Fidx %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Ftenbank %></td>
	<td align=center><%= left(ipkum.Fipkumitem(i).FBankdate,10) %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Fgubun %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Fipkumuser %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Fchulkumsum %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Fipkumsum %></td>
	<td align=center><%= ipkum.Fipkumitem(i).Fremainsum %></td>
	<td align=center>&nbsp;</td>
</tr>
<% next %>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td align=center colspan=10>
	<% if ipkum.HasPreScroll then %>
		<a href="javascript:scrollmove('<%= ipkum.StarScrollPage-1 %>');"><����></a>
	<% else %>
	<% end if %>
	<% for ix = 0 + ipkum.StarScrollPage  to ipkum.StarScrollPage + ipkum.FScrollCount - 1 %>
	<% if (ix > ipkum.FTotalpage) then Exit for %>
	<% if CStr(ix) = CStr(ipkum.FCurrPage) then %>
	<font color="#666666" class="verdana-xsmall"><strong><%= ix %></strong></font>
	<% else %>
	<a href="javascript:scrollmove('<%= ix %>');" class="bb"><font color="#666666"><%= ix %></font></a>
	<% end if %>
	<% next %>
	<% if ipkum.HasNextScroll then %>
	<a href="javascript:scrollmove('<%= ix %>');" class="verdana-xsmall"><����></a>
	<% else %>
	<% end if %></td>
</tr>
</table>
<% set ipkum=nothing %> 

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
