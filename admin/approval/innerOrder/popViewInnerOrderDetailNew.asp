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

dim i, j

idx = requestCheckvar(Request("idx"),32)

if (idx = "") then
	idx = -1
end if

'==============================================================================
dim oinnerordermaster
set oinnerordermaster = New CInnerOrder

oinnerordermaster.FRectIdx = idx
oinnerordermaster.GetInnerOrderOne

dim yyyy, mm
yyyy = Left(oinnerordermaster.FOneItem.FappDate, 4)
mm = Left(oinnerordermaster.FOneItem.FappDate, 7)
mm = Right(mm, 2)

'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FCurrPage = 1
oinnerorder.FPageSize = 500

oinnerorder.FRectIdx = idx

oinnerorder.GetOnOffInnerOrderDetailNew

%>
<script language="javascript">

function jsModifyInnerOrderPercentage(frm) {
	if (frm.innerorderpercentage.value == "") {
		alert("�й������ �Է��ϼ���.");
		return;
	}

	if (frm.innerorderpercentage.value*0 != 0) {
		alert("�й������ ���ڸ� �����մϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "modifyinnerorderpercentage";
		frm.submit();
	}
}

function jsModifyInnerOrderOne(frm) {
	if (confirm("����/�鼼 ���� ��� ���ۼ��˴ϴ�.\n\n���ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "updateOneDetail";
		frm.submit();
	}
}

function popViewDetail(shopid){
	<% if (oinnerordermaster.FOneItem.Fdivcd <> "101") and (oinnerordermaster.FOneItem.Fdivcd <> "102") and (oinnerordermaster.FOneItem.Fdivcd <> "103") and (oinnerordermaster.FOneItem.Fdivcd <> "104") then %>
		alert("�۾����Դϴ�.");
		return;
	<% end if %>

	 var iURI = '/admin/analysis/offgainsumDetail.asp?yyyy1=<%= yyyy %>&mm1=<%= mm %>&shopid=' + shopid;
    var popwin = window.open(iURI,'popViewDetail','width=1000,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<form name="frm" method="post" action="innerOrderDetail_process.asp">
<input type="hidden" name="mode" value="updateOneDetail">
<input type="hidden" name="appDate" value="<%= oinnerordermaster.FOneItem.FappDate %>">
<input type="hidden" name="divcd" value="<%= oinnerordermaster.FOneItem.Fdivcd %>">
<input type="hidden" name="SELLBIZSECTION_CD" value="<%= oinnerordermaster.FOneItem.FSELLBIZSECTION_CD %>">
<input type="hidden" name="BUYBIZSECTION_CD" value="<%= oinnerordermaster.FOneItem.FBUYBIZSECTION_CD %>">
</form>
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30 colspan="2"><b>���ΰŷ� �󼼳���</b></td>
				</tr>
				<tr>
					<td>
						* �¶���/���������� ����ó <font color=red>������</font>�� ���ΰŷ����� ������� �ʴ´�.
					</td>
					<td align="right">
						<input type="button" class="button" value="���ۼ�" onClick="jsModifyInnerOrderOne(frm)">
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="90" align=center>
						����Ʈ(����)
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" align=center>
						����Ʈ(����)��
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						�귣��
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						�ǸŰ���
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="40" align=center>
						������
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						�����
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="40" align=center>
						�й�<br>����
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						���԰���
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="100" align=center>
						�й���
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="60" align=center>
						����<br>����
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						���ΰŷ���
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						�ΰ���
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="75" align=center>
						�հ�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="100" align=center>���</td>
				</tr>
<%

dim makerSupplySum, makerTaxSum, makerTotalSum,totalsellcashSum

makerSupplySum = 0
makerTaxSum = 0
makerTotalSum = 0
totalsellcashSum=0
%>
<%IF oinnerorder.FResultCount > 0 THEN %>
<% for i = 0 to (oinnerorder.FResultCount - 1) %>
	<%
	makerSupplySum = makerSupplySum + oinnerorder.FItemList(i).FmakerSupplySum
	makerTaxSum = makerTaxSum + oinnerorder.FItemList(i).FmakerTaxSum
	makerTotalSum = makerTotalSum + oinnerorder.FItemList(i).FmakerTotalSum

    totalsellcashSum=totalsellcashSum+oinnerorder.FItemList(i).Ftotalsellcash
    
	%>
				<form name="frm<%= i %>" method="post" action="innerOrderDetail_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="idx" value="<%= idx %>">
				<input type="hidden" name="detailidx" value="<%= oinnerorder.FItemList(i).Fdetailidx %>">
				<input type="hidden" name="dealdiv" value="<%= oinnerorder.FItemList(i).Fdealdiv %>">
				<tr>
					<td bgcolor="#FFFFFF" height="30"  align=center>
						<%= oinnerorder.FItemList(i).Fsitename %>
					</td>
					<td bgcolor="#FFFFFF" height="30"  align=center>
						<%= oinnerorder.FItemList(i).Fshopname %>
					</td>
					<td bgcolor="#FFFFFF" align=center><%= oinnerorder.FItemList(i).Fmakerid %></td>
					<td bgcolor="#FFFFFF" align=right>
						<% if (oinnerorder.FItemList(i).Fdealdiv = "2") then response.write "<b>" end if %>
						<%= FormatNumber(oinnerorder.FItemList(i).Ftotalsellcash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= FormatNumber((100 - oinnerorder.FItemList(i).Fsitefee), 0) %>%
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= FormatNumber(oinnerorder.FItemList(i).Ftotalchulgocash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<% if (oinnerorder.FItemList(i).Ftotalsellcash = 0) then %>
							100%
						<% else %>
							<%= CLng(oinnerorder.FItemList(i).Ftotalbuycash / oinnerorder.FItemList(i).Ftotalsellcash * 100) %>%
						<% end if %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<% if (oinnerorder.FItemList(i).Fdealdiv = "1") then response.write "<b>" end if %>
						<%= FormatNumber(oinnerorder.FItemList(i).Ftotalbuycash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= oinnerorder.FItemList(i).GetDealDivName %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<input type="text" class="text" name="innerorderpercentage" size="2" value="<%= oinnerorder.FItemList(i).Finnerorderpercentage %>">%
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<b><%= FormatNumber(oinnerorder.FItemList(i).FmakerSupplySum, 0) %></b>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerTaxSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerTotalSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<input type="button" class="button" value="��ȸ" onClick="popViewDetail('<%= oinnerorder.FItemList(i).Fshopid %>');">
						<input type="button" class="button" value="����" onClick="jsModifyInnerOrderPercentage(frm<%= i %>);">
					</td>
					</form>
				</tr>
<%
	Next
%>
				<tr>
					<td bgcolor="#FFFFFF" height="30">
						�հ�
					</td>
					<td bgcolor="#FFFFFF" colspan=2></td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(totalsellcashSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" colspan=6></td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(makerSupplySum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(makerTaxSum, 0) %>
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
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
