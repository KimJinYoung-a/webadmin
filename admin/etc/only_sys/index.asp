<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script language="javascript">
function goURLiframe(g)
{
	for(var i=1;i<=13;i++)
	{
		document.getElementById("td"+i+"").style.backgroundColor = "#FFFFFF";
	}
	
	if(g == "1")
	{
		document.getElementById("iframe1").src = "event_date_update.asp";
	}
	else if(g == "2")
	{
		document.getElementById("iframe1").src = "userinfo_modify.asp";
	}
	else if(g == "3")
	{
		document.getElementById("iframe1").src = "birthday_coupon.asp";
	}
	else if(g == "4")
	{
		document.getElementById("iframe1").src = "brand_move.asp";
	}
	else if(g == "5")
	{
		document.getElementById("iframe1").src = "tester_date.asp";
	}
	else if(g == "6")
	{
		document.getElementById("iframe1").src = "brand_ordercomment.asp";
	}
	else if(g == "7")
	{
		document.getElementById("iframe1").src = "goodusing.asp";
	}
	else if(g == "8")
	{
		document.getElementById("iframe1").src = "orderlist.asp";
	}
	else if(g == "9")
	{
		document.getElementById("iframe1").src = "giftcard_reg.asp";
	}
	else if(g == "10")
	{
		document.getElementById("iframe1").src = "award_notinclude_item.asp";
	}
	else if(g == "11")
	{
		document.getElementById("iframe1").src = "scm_change_log.asp";
	}
	else if(g == "12")
	{
		document.getElementById("iframe1").src = "dandokgumae.asp";
	}
	else if(g == "13")
	{
		document.getElementById("iframe1").src = "mobile_image_recatch.asp";
	}
	
	document.getElementById("td"+g+"").style.backgroundColor = "#E6E6E6";
}
</script>

<table width="100%" height="100%" border="1">
<tr>
	<td width="15%" valign="top">
		<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td id="td1" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('1')">�̺�Ʈ ��¥ ����</td>
		</tr>
		<tr>
			<td id="td2" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('2')">ȸ�� �̸� ����</td>
		</tr>
		<tr>
			<td id="td3" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('3')">�������� �����߱�</td>
		</tr>
		<tr>
			<td id="td4" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('4')">�귣�� �̵�</td>
		</tr>
		<tr>
			<td id="td5" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('5')">�׽����̺�Ʈ �ı� ��¥����</td>
		</tr>
		<tr>
			<td id="td6" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('6')">�귣�庰 or ��ǰ��<br>�ֹ��� ���ǻ��� ����</td>
		</tr>
		<tr>
			<td id="td7" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('7')">��ǰ�ı� IsUsing ����</td>
		</tr>
		<tr>
			<td id="td8" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('8')">�ֹ�����Ʈ ����</td>
		</tr>
		<tr>
			<td id="td9" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('9')">����Ʈī�� �߱�</td>
		</tr>
		<tr>
			<td id="td10" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('10')">����� ���ܻ�ǰ����</td>
		</tr>
		<tr>
			<td id="td11" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('11')">���� ���� �α�</td>
		</tr>
		<tr>
			<td id="td12" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('12')">�ܵ�����,����������</td>
		</tr>
		<tr>
			<td id="td13" height="20" bgcolor="#FFFFFF" align="center" style="cursor:pointer" onClick="goURLiframe('13')">����� ���̹���<br>�ٽ� ĸ��</td>
		</tr>
		</table>
	</td>
	<td width="85%" valign="top" style="padding:10 0 0 10;">
		<iframe name="iframe1" id="iframe1" src="about:blank" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="yes"></iframe>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->