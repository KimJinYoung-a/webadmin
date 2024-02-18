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
<!-- #include virtual="/lib/classes/approval/innerOrdercls.asp"-->
<%

dim idx

dim i, j

idx = requestCheckvar(Request("idx"),32)

if (idx = "") then
	idx = -1
end if

'==============================================================================
dim oinnerorder
set oinnerorder = New CInnerOrder

oinnerorder.FCurrPage = 1
oinnerorder.FPageSize = 500

oinnerorder.FRectIdx = idx

oinnerorder.GetOnlineInnerOrderDetail

%>
<script language="javascript">

function jsModifyInnerOrderPercentage(frm) {
	if (frm.innerorderpercentage.value == "") {
		alert("분배비율을 입력하세요.");
		return;
	}

	if (frm.innerorderpercentage.value*0 != 0) {
		alert("분배비율은 숫자만 가능합니다.");
		return;
	}

	if (confirm("수정하시겠습니까?") == true) {
		frm.mode.value = "modifyinnerorderpercentage";
		frm.submit();
	}
}

</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >
<tr>
	<td>
		<table width="100%" align="left" cellpadding="1" cellspacing="1" class="a"   border="0" >
		<tr>
			<td>
				<table width="100%" cellpadding="1" cellspacing="1" class="a">
				<tr>
					<td height=30><b>온라인매출(내부거래)</b></td>
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
						사이트
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						매출구분
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						수수료
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						판매가합
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						분배비율
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						브랜드
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						거래액
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						부가세
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>
						합계
					</td>
					<td bgcolor="<%= adminColor("tabletop") %>" align=center>비고</td>
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
				<form name="frm<%= i %>" method="post" action="innerOrderDetail_process.asp">
				<input type="hidden" name="mode" value="">
				<input type="hidden" name="idx" value="<%= idx %>">
				<input type="hidden" name="detailidx" value="<%= oinnerorder.FItemList(i).Fdetailidx %>">
				<tr>
					<td bgcolor="#FFFFFF" height="30"  align=center>
						<%= oinnerorder.FItemList(i).Fsitename %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= oinnerorder.FItemList(i).GetMeachulGubunName %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<%= oinnerorder.FItemList(i).Fsitefee %>%
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).Ftotalsellcash, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<input type="text" class="text" name="innerorderpercentage" size="2" value="<%= oinnerorder.FItemList(i).Finnerorderpercentage %>">%
					</td>
					<td bgcolor="#FFFFFF" align=center>
						<% if (oinnerorder.FItemList(i).Fmakerid = "") then %>
							<%= oinnerorder.FItemList(i).Fshopid %>
						<% else %>
							<%= oinnerorder.FItemList(i).Fmakerid %>
						<% end if %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerSupplySum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerTaxSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" align=right>
						<%= FormatNumber(oinnerorder.FItemList(i).FmakerTotalSum, 0) %>
					</td>
					<td bgcolor="#FFFFFF" width=80 align=center>
						<input type="button" class="button" value="수정" onClick="jsModifyInnerOrderPercentage(frm<%= i %>);">
					</td>
					</form>
				</tr>
<%
	Next
%>
				<tr>
					<td bgcolor="#FFFFFF" height="30">
						합계
					</td>
					<td bgcolor="#FFFFFF" colspan=5></td>
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
					<td colspan="16" align="center">등록된 내역이 없습니다.</td>
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
