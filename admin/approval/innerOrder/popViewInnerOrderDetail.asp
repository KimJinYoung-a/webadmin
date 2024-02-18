<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/innerOrdercls.asp"-->
<%

dim idx
dim groupby

dim i, j

idx = requestCheckvar(Request("idx"),32)
groupby = requestCheckvar(Request("groupby"),32)

if (idx = "") then
	idx = -1
end if

if (groupby = "") then
	groupby = "makerid"
end if

'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FCurrPage = 1
oinnerorder.FPageSize = 500

oinnerorder.FRectIdx = idx
oinnerorder.FRectGroupBy = groupby

oinnerorder.GetInnerOrderDetailNew

%>
<script language="javascript">

function jsRegInsertShopChulg11o(frm) {
	/*
	if (confirm("�ϰ����� �Ͻðڽ��ϱ�?\n\n������ �ð��� �ҿ�˴ϴ�.(5~10��)") == true) {
		frm.mode.value = "reginsertshopchulgo";
		frm.submit();
	}
	*/
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="get">
		<input type="hidden" name="idx" value="<%= idx %>">
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30>
						<input type="radio" name="groupby" value="makerid" onClick="document.frm.submit();" <% if (groupby = "makerid") then %>checked<% end if %> ><b>�귣�庰 �ŷ���</b>
						<input type="radio" name="groupby" value="shopid" onClick="document.frm.submit();" <% if (groupby = "shopid") then %>checked<% end if %> ><b>���庰 �ŷ���</b>
					</td>
				</tr>
				<tr>
					<td>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" align=center>
						�귣��(����)
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						�ŷ���<br>
						<font color=gray>(�ΰ���)</font>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						�հ�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>���</td>
				</tr>
<%

dim makerSupplySum, makerTaxSum, makerTotalSum

makerSupplySum = 0
makerTaxSum = 0
makerTotalSum = 0

%>
<%IF oinnerorder.FResultCount > 0 THEN %>
<% for i = 0 to (oinnerorder.FResultCount - 1) %>
	<%
	makerSupplySum = makerSupplySum + oinnerorder.FItemList(i).FmakerSupplySum
	makerTaxSum = makerTaxSum + oinnerorder.FItemList(i).FmakerTaxSum
	makerTotalSum = makerTotalSum + oinnerorder.FItemList(i).FmakerTotalSum

	%>
				<tr>
					<td bgcolor="#FFFFFF" height="30">
						<% if (oinnerorder.FItemList(i).Fmakerid = "") then %>
							<%= oinnerorder.FItemList(i).Fshopid %>
						<% else %>
							<%= oinnerorder.FItemList(i).Fmakerid %>
						<% end if %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerSupplySum, 0) %><br>
						<font color=gray>(<%= FormatNumber(oinnerorder.FItemList(i).FmakerTaxSum, 0) %>)</font>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerTotalSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center>
						<!--
						<input type="button" class="button" value="��" onClick="jsRegIns11ertShopChulg11o(frm);" disabled>
						-->
					</td>
				</tr>
<%
	Next
%>
				<tr>
					<td bgcolor="#FFFFFF" height="30">
						�հ�
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(makerSupplySum, 0) %><br>
						<font color=gray>(<%= FormatNumber(makerTaxSum, 0) %>)</font>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(makerTotalSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center></td>
				</tr>
<%
	ELSE
%>
				<tr bgcolor="#FFFFFF">
					<td colspan="16" align="center">��ϵ� ������ �����ϴ�.</td>
				</tr>
<%END IF%>

				</table>
			</td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
