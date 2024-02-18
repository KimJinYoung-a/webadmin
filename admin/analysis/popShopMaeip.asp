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
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->
<%

dim yyyymm, shopid, makerid, jungsangubun

yyyymm 			= requestCheckvar(Request("yyyymm"),32)
shopid 			= requestCheckvar(Request("shopid"),32)
makerid 		= requestCheckvar(Request("makerid"),32)
jungsangubun 	= requestCheckvar(Request("jungsangubun"),32)


'// ===========================================================================
dim oshopcostpermeachul
set oshopcostpermeachul = new COffShopCostPerMeachul

oshopcostpermeachul.FRectYYYYMM   = yyyymm
oshopcostpermeachul.FRectShopID   = shopid
oshopcostpermeachul.FRectMakerID  = makerid
oshopcostpermeachul.FRectJungsanGubun = jungsangubun

oshopcostpermeachul.GetOffShopMakerMonthlyMaeip


'==============================================================================
dim i, j

%>
<script language="javascript">

function PopShopMakerMonthlyMaeipDetail(ipchulcode) {
	var viewURL = "";

	<% if (jungsangubun = "B031") then %>
		viewURL = "/admin/newstorage/culgolist.asp?menupos=540&research=on&page=&code=" + ipchulcode;
	<% elseif (jungsangubun = "B022") then %>
		viewURL = "/common/offshop/shop_ipchuldetail.asp?menupos=196&idx=" + ipchulcode;
	<% elseif (jungsangubun = "B012") or (jungsangubun = "B011") then %>
		viewURL = "/admin/offupchejungsan/off_jungsanlist.asp?menupos=926&makerid=<%= makerid %>";
	<% else %>
		viewURL = "";
	<% end if %>

	if (viewURL == "") {
		alert("�۾����Դϴ�.");
		return;
	}

	var popwin = window.open(viewURL,'PopShopMakerMonthlyMaeipDetail','width=1200, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<form name="frm" method="post" action="popRegInnerOrderByMonth_process11.asp">
		<input type="hidden" name="mode" value="">
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>��ǰ�� ���� ���Ծ�</b></td>
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
					<td bgcolor="<%= adminColor("tabletop") %>" height="30" width="80" align=center>
						����
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="100" align=center>
						�����ڵ�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						�귣��
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80" align=center>
						��ǰ�ڵ�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						��ǰ��<br><font color=blue>[�ɼǸ�]</font>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="70" align=center>
						�������
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="70" align=center>
						������԰�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="40" align=center>
						����
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>"width="70"  align=center>
						���Ծ�
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>���</td>
				</tr>
<%

dim buycashsum

buycashsum = 0

%>
<% for i = 0 to (oshopcostpermeachul.FResultCount - 1) %>
	<%
	buycashsum = buycashsum + oshopcostpermeachul.FItemList(i).Fbuycash * oshopcostpermeachul.FItemList(i).Fitemno
	%>
				<tr>
					<td bgcolor="#FFFFFF" height="30" align=center>
						<%= oshopcostpermeachul.FItemList(i).Fshopid %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<a href="javascript:PopShopMakerMonthlyMaeipDetail('<%= oshopcostpermeachul.FItemList(i).Fipchulcode %>')">
							<%= oshopcostpermeachul.FItemList(i).Fipchulcode %>
						</a>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= oshopcostpermeachul.FItemList(i).Fmakerid %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= oshopcostpermeachul.FItemList(i).GetBarcode %>
					</td>
					<td bgcolor="#FFFFFF">
						<%= oshopcostpermeachul.FItemList(i).Fitemname %><br><font color=blue>[<%= oshopcostpermeachul.FItemList(i).Fitemoptionname %>]</font>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fsuplycash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fbuycash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oshopcostpermeachul.FItemList(i).Fitemno, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber((oshopcostpermeachul.FItemList(i).Fbuycash * oshopcostpermeachul.FItemList(i).Fitemno), 0) %>
					</td>
					<td bgcolor="#FFFFFF">
					</td>
				</tr>
<%
	Next
%>
				<tr>
					<td bgcolor="#FFFFFF" height="30" colspan="8" align="right">
						�հ�
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(buycashsum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center></td>
				</tr>
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
