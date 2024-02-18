<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim vQuery, cItem, vItemID, arrList, intLoop, vIsDandok, vIsSunChak
	vItemID = Replace(requestCheckVar(Request("itemid"),300)," ","")
	vIsDandok = NullFillWith(requestCheckVar(Request("dandok"),1),"o")
	vIsSunChak = requestCheckVar(Request("sunchak"),1)
	
	If vItemID <> "" Then
		Set cItem = new cOnlySys
		cItem.FItemID = vItemID

		arrList = cItem.fnItemDetail
		Set cItem = Nothing
	End IF
	
	vQuery = vQuery & "-- �ܵ����Ż�ǰ(�������� ��) ����" & vbCrLf
	vQuery = vQuery & "select reserveItemTp, availPayType, * from db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "-- �ű��� ���� (0-�Ϲ�, 1-�ܵ�����)" & vbCrLf
	vQuery = vQuery & "-- ����������(�ǽð�/���) ��ǰ ���� (availPayType- 8:Just1Day ,9:����������)" & vbCrLf
	vQuery = vQuery & "--update db_item.dbo.tbl_item" & vbCrLf
	vQuery = vQuery & "set reserveItemTp='1'" & vbCrLf
	vQuery = vQuery & "--, availPayType='9'" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf
%>

<script language="javascript">
function jsItemSearch()
{
	if(frm1.itemid.value == "")
	{
		alert("��ǰ�ڵ� �ʼ��Դϴ�.");
		frm1.itemid.focus();
		return;
	}
	frm1.submit();
}
function jsItemUpdate()
{
	if(frm1.itemid.value == "")
	{
		alert("itemid���� �ʿ��մϴ�.");
		frm1.itemid.focus();
		return;
	}
	
	if(confirm("�̴�� �����Ͻðڽ��ϱ�?") == true) {
		frm1.method = "post";
		frm1.action = "dandokgumae_proc.asp";
		frm1.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr>
	<td>
		<form name="frm1" action="<%=CurrURL%>" method="get">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				ItemID : <input type="text" name="itemid" value="<%=vItemID%>" size="100">&nbsp;
				<input type="button" class="button" value="�� ��" onClick="jsItemSearch()"> * ,��ǥ�� ������ �Է�
			</td>
		</tr>
		<% If vItemID <> "" Then %>
			<tr>
				<td><br>
				<input type="checkbox" name="dandok" value="o" <%=CHKIIF(vIsDandok="o","checked","")%>>�ܵ����ż���(reserveItemTp=1)&nbsp;&nbsp;&nbsp;
				<input type="checkbox" name="sunchak" value="o" <%=CHKIIF(vIsSunChak="o","checked","")%>>����������(�ǽð�/���)����(availPayType=9)
				<input type="button" value="�ٷκ����ϱ�" onClick="jsItemUpdate()"></td>
			</tr>
		<% End If %>
		</table>
		</form>
		<% IF isArray(arrList) THEN %>
		<br>
		[db_item].[dbo].[tbl_Item]
		<table border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>ItemID</td>
			<td>ItemName</td>
		  	<td>reserveItemTp</td>
		  	<td>availPayType</td>
		  	<td>RegDate</td>
		  	<td>lastupdate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><%=arrList(2,intLoop)%></td>
					<td><%=arrList(3,intLoop)%></td>
					<td><%=arrList(4,intLoop)%></td>
					<td><%=arrList(5,intLoop)%></td>
				</tr>
			<% Next %>
		</table>
		<% End If %>
	</td>
</tr>
</table>


<% If vItemID <> "" Then %>
<br><br>* ��������<br>
<textarea name="" cols="100" rows="15"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->