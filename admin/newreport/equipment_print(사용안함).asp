<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	2010�� 01�� 06�� �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->
<%
' ���� ����
dim page, jangbi, sayoug, idx, user, usingIp , code 
dim ipgocheck, yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate, ipcheck, useip
	page = request("page")
	if page="" then page=1
	jangbi = Request("jangbi")		'���˻��� �ʿ��Ѻ���
	sayoug = Request("sayoug")		'��뱸�п� �ʿ��� ����
	user = Request("user")			'����� �˻��� �ʿ��Ѻ���
	idx = Request("idx")			'������ �ε��� ����
	code = request("code")			'����ڵ� �˻��� �ʿ��� ����
	ipcheck = request("ipcheck")		'ip�˻��� �ʿ��� ����

	' �԰��� �˻��� �ʿ��� ���� ����
	ipgocheck = request("ipgocheck")
	yyyy1 = request("yyyy1")
	yyyy2 = request("yyyy2")
	mm1	  = request("mm1")
	mm2	  = request("mm2")
	dd1	  = request("dd1")
	dd2	  = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (ipcheck <> "") then ipcheck = "on"

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oequip,i				'class ����
set oequip = new CEquipment		'class���� ����
	oequip.FPageSize = 50
	oequip.FCurrPage = page
	oequip.FRectJangbi = jangbi
	oequip.FRectSayoug = sayoug
	oequip.FRectUser = user
	oequip.FRectIp = ipcheck
	oequip.Fequip_code = code
		if ipgocheck = "on" then
			oequip.FRectBuyDateDtStart = fromDate
			oequip.FRectBuyDateDtEnd = toDate
		end if
	
	oequip.getEquipmentList		'class �Լ� ����

'���� ����
Dim equip_gubun, part_code
	equip_gubun = Request("equip_gubun")	'��񱸺�
	part_code = Request("part_code")		'��뱸��

dim totalcurrsum,totaljasan, Alltotaljasan, getAllCurrentValue
	totalcurrsum = 0	'���� �������� ���Ű��� �հ賻�� ���� ����.
	totaljasan	 = 0	'���� �������� �ڻ갡ġ�� �հ賻�� ���� ����.
	Alltotaljasan = 0
%>

<script language="javascript">

	window.onload = function regprint(){
		window.print();
		self.close();
	}
	
</script>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" border=1 cellspacing="1" bordercolor="black">
	<% if oequip.FResultCount > 0 then %>
	<tr height="25" >
		<td colspan="10">
			�˻���� : <b><%= oequip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oequip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" >
    	<td >���<br>�ڵ�</td>
    	<td >�����<br>�̸�</td>
    	<td >���</td>
    	<td >���<br>����</td>
    	<td >���<br>����</td>
    	<td >��ǰ��</td>
    	<td >������</td>
    	<td >����<br>����</td>    	
    	<td >���Ű�</td>
    	<td >�ڻ�<br>��ġ</td>
    </tr>
<% for i=0 to oequip.FResultCount - 1 %>
	<%
	totalcurrsum = totalcurrsum + oequip.FItemList(i).Fbuy_sum 		'�������� �հ踦 ���ϱ� ���ؼ� ���� �������� ���Ű�(Fbuy_sum)�� ��� (fot~next�� loop�� ������)���ؼ� totalcurrsum ������ ����
	totaljasan	 = totaljasan + oequip.FItemList(i).GetCurrentvalue	'�������� �ڻ갡ġ�� ���ϱ� ���� ���� �������� �ڻ갡ġ(GetCurrentvalue)�� ��� for ������ ������ ���ϰ� totaljasan ������ ����
	%>
	<form name=frm_<%= i %> method="post" action="frmdel.asp">	<!-- for�� �ȿ��� i���� ������ ����-->
	<input type="hidden" name="idx" value="<%= oequip.FItemList(i).Fidx %>">
	<input type="hidden" name="ssBctId" value="<%= session("ssBctId")%>">
    <tr align="center" >
		<td><%= oequip.FItemList(i).Fequip_code %></td>
		<td><%= oequip.FItemList(i).FusinguserName %>&nbsp;&nbsp;<%= (oequip.FItemList(i).Fusinguserid) %></td>
		<td><%= oequip.FItemList(i).Fdetail_quality1 %><br><%= oequip.FItemList(i).Fdetail_quality2 %></td>
		<td><%= oequip.FItemList(i).Fequip_gubun_name %></td>
		<td><%= oequip.FItemList(i).Fpart_code_name %></td>
		<td><%= oequip.FItemList(i).Fequip_name %></td>
		<td align="center"><%= oequip.FItemList(i).Fbuy_date %></td>
		<td align="right"><%= formatNumber(oequip.FItemList(i).fwonga_cost,0) %></td>		
		<td align="right"><%= formatNumber(oequip.FItemList(i).Fbuy_sum,0) %></td>
		<td align="right">
			<% if oequip.FItemList(i).getCurrentValue<>"" then %>
				<font color="#EE3333"><%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%></font>
			<% else %>
				<%=formatNumber(oequip.FItemList(i).GetCurrentvalue,0)%>
			<% end if %>
		</td>				
    </tr>   
	</form>
	<% next %>
	
	<tr >
		<td align="center" colspan=7>�Ѱ�</td>
		<td align="right"><!-- <%= formatNumber(oequip.FItemList(0).Getallcurrentvalue,0) %> --></td>
		<td align="right"><%= formatNumber(totalcurrsum,0) %></td>
		<td align="right"><font color="#EE3333"><%= formatNumber(totaljasan,0) %></font></td>
	
	</tr>
		

	<% else %>
		<tr >
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
</table>
</body>
</html>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->