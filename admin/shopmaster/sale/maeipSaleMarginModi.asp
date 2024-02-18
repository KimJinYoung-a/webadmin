<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemMaeipSaleMarginShareCls.asp"-->
<%

dim idx
idx     = requestCheckvar(request("idx"),32)


'==============================================================================
dim oMaster, oDetail

set oMaster = new CItemMaeipSaleMarginShare
oMaster.FRectIdx         = idx
oMaster.GetMasterOne

set oDetail = new CItemMaeipSaleMarginShare
oDetail.FPageSize 		= 500
oDetail.FRectIdx    	= idx
if (idx <> "") then
	oDetail.GetDetailList
end if


dim mode, i

mode = "modi"
if oMaster.FResultCount < 1 then
	mode = "ins"
end if

%>
<script language="javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jsSubmitSave() {
	var frm = document.frm;

	if (frm.makerid.value == "") {
		alert("브랜드를 입력하세요.");
		return;
	}

	if (frm.saleCode.value == "") {
		alert("할인코드를 입력하세요.");
		return;
	}

	if (frm.saleCode.value*0 != 0) {
		alert("할인코드는 숫자만 입력가능 합니다.");
		return;
	}

	if ((frm.startDate.value == "") || (frm.endDate.value == "")) {
		alert("기간을 입력하세요.");
		return;
	}

	if (frm.defaultMargin.value == "") {
		alert("기본마진을 입력하세요.");
		return;
	}

	if (frm.defaultMargin.value*0 != 0) {
		alert("기본마진은 숫자만 입력가능 합니다.");
		return;
	}

	if (frm.saleMargin.value == "") {
		alert("할인마진을 입력하세요.");
		return;
	}

	if (frm.saleMargin.value*0 != 0) {
		alert("할인마진은 숫자만 입력가능 합니다.");
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
		frm.submit();
	}
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="maeipSaleMargin_process.asp" onSubmit="return false;">
	<input type="hidden" name="mode" value="<%= mode %>">
	<input type="hidden" name="idx" value="<%= oMaster.FOneItem.Fidx %>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr height="25">
		<td width="7%" bgcolor="<%= adminColor("tabletop") %>" align="center">IDX</td>
		<td width="43%" bgcolor="#FFFFFF"><%= oMaster.FOneItem.Fidx %></td>
		<td width="7%" bgcolor="<%= adminColor("tabletop") %>" align="center"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">할인코드</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="saleCode" size="14" maxlength="64" value="<%= oMaster.FOneItem.FsaleCode %>">
			<!--
			<input type="button" class="button" value="가져오기" onClick="alert('aa');">
			-->
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">브랜드</td>
		<td bgcolor="#FFFFFF">
			<%	drawSelectBoxDesignerWithName "makerid", oMaster.FOneItem.Fmakerid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">기간(결제일자)</td>
		<td bgcolor="#FFFFFF">
			시작일 : <input type="text" name="startDate" size="13" onClick="jsPopCal('startDate');" style="cursor:hand;" value="<%= oMaster.FOneItem.FstartDate %>"  class="text">
			~
			종료일 : <input type="text" name="endDate" size="13" onClick="jsPopCal('endDate');" style="cursor:hand;" value="<%= oMaster.FOneItem.FendDate %>"  class="text">
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">매출구분</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="meachulGubun">
				<!-- 출고일 기준이 더 깔끔하다., skyer9, 2018-03-21
				<option value="1" <%= CHKIIF(oMaster.FOneItem.FmeachulGubun="1", "selected", "")%>>결제일 기준</option>
				-->
				<option value="2" <%= CHKIIF(oMaster.FOneItem.FmeachulGubun="2", "selected", "")%>>출고일 기준</option>
			</select>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">
			기본마진
		</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="defaultMargin" size="4" maxlength="64" value="<%= oMaster.FOneItem.FdefaultMargin %>">% (소비자가 대비)
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">
			할인마진
		</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="saleMargin" size="4" maxlength="64" value="<%= oMaster.FOneItem.FsaleMargin %>">% (할인가 대비)
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">사용여부</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="useyn">
				<option value="Y" <%= CHKIIF(oMaster.FOneItem.Fuseyn="Y", "selected", "")%>>사용</option>
				<option value="N" <%= CHKIIF(oMaster.FOneItem.Fuseyn="N", "selected", "")%>>사용안함</option>
			</select>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">등록자</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Freguserid %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">등록일</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Fregdate %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">최종수정</td>
		<td bgcolor="#FFFFFF">
			<%= oMaster.FOneItem.Flastupdate %>
		</td>
	</tr>
	<tr height="50">
		<td bgcolor="#FFFFFF" colspan="4" align="center">
			<input type="button" class="button" value="저장하기" onClick="jsSubmitSave();">
			<input type="button" class="button" value="취소하기" onClick="location.href='maeipSaleMarginList.asp?menupos=<%= menupos %>';">
		</td>
	</tr>
	</form>
</table>

<p />

<% if (mode = "modi") then %>
※ 매입상품만 적용가능합니다.<br />
※ 마진구분은 텐바이텐부담만 등록 가능합니다.(이미 매입한 상품이라 마진쉐어가 안되고, 판매장려금을 설정합니다.)<br />
※ 할인코드(<%= oMaster.FOneItem.FsaleCode %>) 에 등록된 브랜드(<%= oMaster.FOneItem.Fmakerid %>) 상품중 <font color="red">기본마진 <%= oMaster.FOneItem.FdefaultMargin %>%</font> 이고, <font color="red">기본매입가:할인매입가 동일한</font> 상품만 추가됩니다.

<p />

<% end if %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
		<td colspan="17" align="left">검색결과 : <b><%= oDetail.FResultCount %> / <%= oDetail.FTotalCount %></b></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="80">상품코드</td>
		<td align="center" width="55">이미지</td>
		<td align="center">브랜드</td>
		<td align="center">상품명</td>
		<td align="center" width="60">계약<br>구분</td>

		<td align="center" width="80">소비자가</td>
		<td align="center" width="80"><b>기본<br />매입가</b></td>
		<td align="center" width="80">기본<br />마진율</td>

		<td align="center" width="80">할인가</td>
		<td align="center" width="80">할인률</td>
		<td align="center" width="80"><b>할인적용<br />매입가</b></td>
		<td align="center" width="80">할인적용<br />마진율</td>
		<td align="center" width="80"><b>판매<br />장려금</b></td>
		<td>비고</td>
	</tr>
<% if oDetail.FresultCount > 0 then %>
   	<% for i=0 to oDetail.FresultCount-1 %>
	<tr align="center">
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fitemid %></td>
		<td bgcolor="#FFFFFF">
			<img src="<%= oDetail.FItemList(i).Fsmallimage %>">
		</td>
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fmakerid %></td>
		<td bgcolor="#FFFFFF"><%= oDetail.FItemList(i).Fitemname %></td>
		<td bgcolor="#FFFFFF">
			<%= oDetail.FItemList(i).getMwDivName %>
			<% if (oDetail.FItemList(i).FmwDiv <> oDetail.FItemList(i).Fcurrmwdiv) then %>
			<br />(현:<%= oDetail.FItemList(i).getCurrMwDivName %>)
			<% end if %>
		</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oDetail.FItemList(i).Forgprice,0) %></td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber(oDetail.FItemList(i).ForgBuyCash,0) %></b></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Forgprice - oDetail.FItemList(i).ForgBuyCash) / oDetail.FItemList(i).Forgprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><%= FormatNumber(oDetail.FItemList(i).Fsaleprice,0) %></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Forgprice - oDetail.FItemList(i).Fsaleprice) / oDetail.FItemList(i).Forgprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber(oDetail.FItemList(i).FsaleBuyCash,0) %></b></td>
		<td bgcolor="#FFFFFF">
			<%= (oDetail.FItemList(i).Fsaleprice - oDetail.FItemList(i).FsaleBuyCash) / oDetail.FItemList(i).Fsaleprice * 100 %>%
		</td>
		<td bgcolor="#FFFFFF"><b><%= FormatNumber((oDetail.FItemList(i).ForgBuyCash - oDetail.FItemList(i).FsaleBuyCash),0) %></b></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
	<% next %>
<% end if %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
