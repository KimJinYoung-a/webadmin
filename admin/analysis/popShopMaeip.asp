<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<%
'###########################################################
' Description : 결제요청서 등록
' History : 2011.03.14 정윤정  생성
' 0 요청/1 진행중/ 5 반려/7 승인/ 9 완료
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
		alert("작업중입니다.");
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
					<td height=30><b>상품별 매장 매입액</b></td>
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
						매장
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="100" align=center>
						입출코드
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						브랜드
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="80" align=center>
						상품코드
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						상품명<br><font color=blue>[옵션명]</font>
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="70" align=center>
						매장출고가
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="70" align=center>
						본사매입가
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" width="40" align=center>
						수량
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>"width="70"  align=center>
						매입액
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>비고</td>
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
						합계
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
