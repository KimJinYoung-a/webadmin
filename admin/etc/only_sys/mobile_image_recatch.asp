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
	Dim vQuery, cItem, vItemID, arrList, intLoop
	vItemID = Replace(requestCheckVar(Request("itemid"),300)," ","")
	
	If vItemID <> "" Then
		Set cItem = new cOnlySys
		cItem.FItemID = vItemID

		arrList = cItem.fnItemDetail
		Set cItem = Nothing
	End IF
	
	
	vQuery = vQuery & "select * from db_etcmall.dbo.tbl_outmall_API_Que" & vbCrLf
	vQuery = vQuery & "where itemid in (" & vItemID & ")" & vbCrLf & vbCrLf
	
	vQuery = vQuery & "--insert into db_etcmall.dbo.tbl_outmall_API_Que" & vbCrLf
	vQuery = vQuery & "select 'appDTL','EDIT',itemid,1100,GETDATE(),NULL,NULL,NULL,NULL,'" & session("ssBctId") & "'" & vbCrLf
	vQuery = vQuery & "from db_item.dbo.tbl_item" & vbCrLf
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
		frm1.action = "mobile_image_recatch_proc.asp";
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
				ItemID : <input type="text" name="itemid" value="<%=vItemID%>" maxlength="32" size="100">&nbsp;
				<input type="button" class="button" value="�� ��" onClick="jsItemSearch()"> * ,��ǥ�� ������ �Է�
			</td>
		</tr>
		<% If vItemID <> "" Then %>
			<tr>
				<td>
				<input type="button" value="�ٽ�ĸ���ϱ�" onClick="jsItemUpdate()"></td>
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
		  	<td>PC��ũ</td>
		  	<td>������ũ</td>
		  	<td>��۸�ũ</td>
		  	<td>lastupdate</td>
		</tr>
			<% For intLoop =0 To UBound(arrList,2) %>
				<tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
					<td><%=arrList(0,intLoop)%></td>
					<td><%=arrList(1,intLoop)%></td>
					<td><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[PC��ũ]</a></td>
					<td><a href="http://m.10x10.co.kr/category/category_itemprd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[������ũ]</a></td>
					<td><a href="http://m.10x10.co.kr/apps/appCom/wish/web2014/category/category_itemprd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank">[��۸�ũ]</a></td>
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