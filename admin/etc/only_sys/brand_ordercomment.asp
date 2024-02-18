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
	Dim cBrandOrder, vQuery, vMakerID, vItemID, vMoveItemCnt, vChange
	vMakerID = requestCheckVar(Request("makerid"),100)
	vItemID = Request("itemid")
	
	If vMakerID <> "" OR vItemID <> "" Then
		Set cBrandOrder = new cOnlySys
		cBrandOrder.FMakerID = vMakerID
		cBrandOrder.FItemID = vItemID
		cBrandOrder.fnBrandOrderCont
		
		vMoveItemCnt = cBrandOrder.FMoveItemCnt
		Set cBrandOrder = Nothing
	End IF
	'wimax
	
	vQuery = ""
	vQuery = vQuery & "select * from db_item.dbo.tbl_item"
	If vMakerID <> "" Then
		vQuery = vQuery & " where makerid = '" & vMakerID & "'" & vbCrLf & vbCrLf
	End If
	If vItemID <> "" Then
		vQuery = vQuery & " where itemid IN(" & vItemID & ")" & vbCrLf & vbCrLf
	End If
	vQuery = vQuery & "select ordercomment, * from db_item.dbo.tbl_item_Contents" & vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "where itemid in (" & vbCrLf
		vQuery = vQuery & "	select itemid" & vbCrLf
		vQuery = vQuery & "	from db_item.dbo.tbl_item" & vbCrLf
		vQuery = vQuery & "	where makerid = '" & vMakerID & "'" & vbCrLf
		vQuery = vQuery & ")" & vbCrLf & vbCrLf
	End If
	If vItemID <> "" Then
		vQuery = vQuery & "where itemid IN(" & vItemID & ")" & vbCrLf & vbCrLf
	End If
	vQuery = vQuery & "--update db_item.dbo.tbl_item_Contents set" & vbCrLf
	vQuery = vQuery & "ordercomment = ''" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vbCrLf
	If vMakerID <> "" Then
		vQuery = vQuery & "	select itemid" & vbCrLf
		vQuery = vQuery & "	from db_item.dbo.tbl_item" & vbCrLf
		vQuery = vQuery & "	where makerid = '" & vMakerID & "'" & vbCrLf
	End If
	If vItemID <> "" Then
		vQuery = vQuery & "" & vItemID & "" & vbCrLf
	End If
	vQuery = vQuery & ")" & vbCrLf
%>

<script language="javascript">
function jsbrandSearch()
{
	if(frm1.makerid.value == "" && frm1.itemid.value == "")
	{
		alert("������ �귣��(WHERE makerid = '')\n�Ǵ� ������ ��ǰ�ڵ�(WHERE itemid IN (''))��\n�Է��ϼ���.");
		return;
	}
	if(frm1.makerid.value != "" && frm1.itemid.value != "")
	{
		alert("������ �귣��(WHERE makerid = '')\n�Ǵ� ������ ��ǰ�ڵ�(WHERE itemid IN (''))��\n�ϳ��� �Է��ϼ���.");
		return;
	}
	frm1.submit();
}
function jsBrandUpdate()
{
	if(frm1.makerid.value == "" && frm1.itemid.value == "")
	{
		alert("������ �귣��(WHERE makerid = '')\n�Ǵ� ������ ��ǰ�ڵ�(WHERE itemid IN (''))��\n�Է��ϼ���.");
		return;
	}
	if(frm1.makerid.value != "" && frm1.itemid.value != "")
	{
		alert("������ �귣��(WHERE makerid = '')\n�Ǵ� ������ ��ǰ�ڵ�(WHERE itemid IN (''))��\n�ϳ��� �Է��ϼ���.");
		return;
	}
	if(frm1.comment.value == "")
	{
		alert("������ �귣���� �ֹ��� ���ǻ��� ������ �Է��ϼ���.");
		frm1.comment.focus()
		return;
	}

	if(confirm("�̴�� �����Ͻðڽ��ϱ�?") == true) {
		frm1.method = "post";
		frm1.action = "brand_ordercomment_proc.asp";
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
			<td colspan="2">
				<font color="blue">1</font>. ������ �귣��(WHERE makerid = '') : <input type="text" name="makerid" value="<%=vMakerID%>" maxlength="50">&nbsp;�Ǵ�(�� �� �ϳ���)<br>
				<font color="blue">2</font>. ������ ��ǰ�ڵ�(WHERE itemid IN ('')) : <textarea name="itemid" cols="30" rows="7"><%=vItemID%></textarea><br>
				�� �Է� �� (123456 �Ǵ� 123456,234567,345678 �Ǵ� 123456, 234567, 345678)<br><br>
				<input type="button" class="button" value="��      ��" onClick="jsbrandSearch()">
			</td>
		</tr>
		</table>
		<% If vMakerID <> "" OR vItemID <> "" Then %>
		<br><br>
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td colspan="2">
				�ֹ��� ���ǻ��� ���� : <textarea name="comment" cols="30" rows="7"></textarea>
				<% If vChange = "o" Then %><font color="red">����Ϸ�.</font><% Else %>�� <%=vMoveItemCnt%>�� ��ǰ<% End If %> <input type="button" value="�ٷκ����ϱ�" onClick="jsBrandUpdate()">
			</td>
		</tr>
		</table>
		<% End If %>
		</form>
	</td>
</tr>
</table>

<% If vMakerID <> "" OR vItemID <> "" Then %>
<br><br>* ��������<br>
<textarea name="" cols="100" rows="17"><%=vQuery%></textarea>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->