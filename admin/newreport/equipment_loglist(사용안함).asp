<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��� �ڻ� ����Ʈ
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->

<%
dim page, jangbi, sayoug, idx, user, usingIp , code ,equip_name ,manufacture_company
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, ipcheck, useip
dim totalcurrsum,totaljasan, Alltotaljasan, getAllCurrentValue ,equip_gubun, part_code
	page = request("page")
	if page="" then page=1
	jangbi = Request("jangbi")		'���˻��� �ʿ��Ѻ���
	sayoug = Request("sayoug")		'��뱸�п� �ʿ��� ����
	user = Request("user")			'����� �˻��� �ʿ��Ѻ���
	idx = Request("idx")			'������ �ε��� ����
	code = request("code")			'����ڵ� �˻��� �ʿ��� ����
	ipcheck = request("ipcheck")		'ip�˻��� �ʿ��� ����
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")
	equip_gubun = Request("equip_gubun")
	part_code = Request("part_code")
	equip_name = Request("equip_name")
	manufacture_company = Request("manufacture_company")
			
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (ipcheck <> "") then ipcheck = "on"

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oequip,i
set oequip = new CEquipment
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectJangbi = jangbi
	oequip.FRectSayoug = sayoug
	oequip.FRectUser = user
	oequip.FRectIp = ipcheck
	oequip.Fequip_code = code
	oequip.frectequip_name = equip_name
	oequip.frectmanufacture_company = manufacture_company
	
	if ipgocheck = "on" then
		oequip.FRectBuyDateDtStart = fromDate
		oequip.FRectBuyDateDtEnd = toDate
	end if
	
	oequip.getEquipmentlogList

totalcurrsum = 0
totaljasan	 = 0
Alltotaljasan = 0
%>

<script language='javascript'>

function NextPage(page){
	frm.page.value= page;
	frm.submit();
}

//������ üũ
function EnDisabledDateBox(comp){
	document.frm.yyyy1.disabled = !comp.checked;
	document.frm.yyyy2.disabled = !comp.checked;
	document.frm.mm1.disabled = !comp.checked;
	document.frm.mm2.disabled = !comp.checked;
	document.frm.dd1.disabled = !comp.checked;
	document.frm.dd2.disabled = !comp.checked;
}

function UseIpCheck(comp){
	//document.frm.ipcheck.disabled = comp.checked;
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<input type=checkbox name="ipgocheck" value="on" <% if ipgocheck="on" then  response.write "checked" %> onclick="EnDisabledDateBox(this)">	
		������<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %><br>
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="NextPage('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		����� : <% drawpartneruser "user", user ,"" %>
		��񱸺� : <% DrawEquipMentGubun "10","jangbi",jangbi ,""%>
		��뱸�� : <% DrawEquipMentGubun "20","sayoug",sayoug ,"" %>
	</td>	
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<!--<input type="checkbox" name="ipcheck" value="on" <%if ipcheck="on" then response.write "checked" %>>��� IP-->
		����ڵ� : <input type="text" name="code" value="<%=code%>">
		��ǰ�� : <input type="text" name="equip_name" value="<%=equip_name%>">
		������ : <input type="text" name="manufacture_company" value="<%=manufacture_company%>">		
	</td>	
</tr>
</form>
</table>
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
</form>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oequip.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oequip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����ڵ�</td>
	<td>�����</td>
	<td>���</td>
	<td>���<br>����</td>
	<td>���<br>����</td>
	<td>��ǰ��</td>
	<td>������</td>
	<td>����<br>����</td>    	
	<td>���Ű�</td>
	<td>�ڻ�<br>��ġ</td>
	<td>����</td>
</tr>
<% if oequip.FResultCount > 0 then %>
<% for i=0 to oequip.FResultCount - 1 %>
<%
totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum
totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue
%>
<form name=frm_<%= i %> method="post" action="frmdel.asp">	<!-- for�� �ȿ��� i���� ������ ����-->
<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
<input type="hidden" name="ssBctId" value="<%= session("ssBctId")%>">
<tr align="center" bgcolor="#FFFFFF" onMouseOver= this.style.background='f1f1f1'; onMouseOut=this.style.background='#ffffff';>
	<td width=130>
		<%= oequip.FItemList(i).Fequip_code %>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).FusinguserName %>
		<% if oequip.FItemList(i).fstatediv <> "Y" then %>
			<font color="red">[���]</font>
		<% end if %>
		
		<% if oequip.FItemList(i).Fusinguserid <> "" then %>
			<Br><%= oequip.FItemList(i).Fusinguserid %>
		<% end if %>
	</td>
	<td>
		<%= oequip.FItemList(i).Fdetail_quality1 %><br><%= oequip.FItemList(i).Fdetail_quality2 %>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).Fequip_gubun_name %>
	</td>
	<td width=100>
		<%= oequip.FItemList(i).Fpart_code_name %>
	</td>
	<td>
		<%= oequip.FItemList(i).Fequip_name %>
	</td>
	<td width=80>
		<%= oequip.FItemList(i).Fbuy_date %>
	</td>
	<td align="right" width=70>
		<%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %>
	</td>		
	<td align="right" width=70>
		<%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %>
	</td>
	<td align="right" width=70>
		<% if oequip.FItemList(i).getCurrentValue<>"" then %>
			<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
		<% else %>
			<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
		<% end if %>
	</td>
	<td align="center" width=60>
		<%= oequip.FItemList(i).fdel_id %>
		<br><%= oequip.FItemList(i).fdel_date %>
	</td>
</tr>   
</form>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan=7>�Ѱ�</td>
	<td align="right"><!-- <%= formatNumber(oequip.FItemList(0).Getallcurrentvalue,0) %> --></td>
	<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
	<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
	<td align="right" colspan=3><!-- ���к� Total : <%= formatNumber(oequip.FTotalSum,0) %> --></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
    	<% if oequip.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oequip.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oequip.StartScrollPage to oequip.FScrollCount + oequip.StartScrollPage - 1 %>
			<% if i>oequip.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oequip.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>

</table>

<%
	set oequip = Nothing
%>

<script language='javascript'>
	EnDisabledDateBox(document.frm.ipgocheck);
	//UseIpCheck(document.frm.ipcheck);
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->