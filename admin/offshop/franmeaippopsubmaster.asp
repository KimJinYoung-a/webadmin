<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 정산
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim idx
	idx = requestCheckVar(request("idx"),10)

dim ofranchulgomaster
set ofranchulgomaster = new CFranjungsan
ofranchulgomaster.FRectidx = idx
ofranchulgomaster.getOneFranJungsan


dim ofranchulgojungsan

set ofranchulgojungsan = new CFranjungsan
ofranchulgojungsan.FPageSize=200
ofranchulgojungsan.FRectIDx = idx
ofranchulgojungsan.getFranMaeipSubmasterList

dim oCCartoonBoxMasterItem

set oCCartoonBoxMasterItem = new CCartoonBoxMasterItem

oCCartoonBoxMasterItem.Fdelivermethod = ofranchulgomaster.FOneItem.Fdelivermethod

dim i

dim totalsellcash,totalbuycash,totalsuplycash,totalorgsellcash

dim mode

if (ofranchulgomaster.FOneItem.Fworkidx <> "") then
	mode = "updateworkidx"
else
	mode = "insertworkidx"
end if

%>
<script language='javascript'>
function popSubdetailEdit(iid,itopid){
	var popwin = window.open('franmeaippopsubdetail.asp?idx=' + iid + '&topidx=' + itopid,'franmeaippopsubdetail','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popEtcAdd(topidx,shopid){
	if ("<%=ofranchulgomaster.FOneItem.FstateCd%>" >= "4")
	{
		alert("계산서 발행 이후에는 기타내역추가 할 수 없습니다.")
		return;
	}

	var popwin = window.open('franetcjungsanadd.asp?topidx=' + topidx + '&shopid=' + shopid,'franetcjungsanadd','width=600, height=150, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function UpdateWorkidx(frm) {
	if (CheckBox(frm) == true) {
		if (confirm('저장하시겠습니까?') == true) {
			frm.submit();
		}
	}
}

function popAddDeliverPay(frm) {
	if (frm == undefined) {
		alert("해외출고에만 EMS비용을 추가할 수 있습니다.");
		return;
	}

	if (frm.workidx.value == "") {
		alert("먼저 관련작업을 입력하세요.");
		return;
	}

	if (CheckBox(frm) == true) {
		if (confirm('EMS운송비용을 추가하시겠습니까?') == true) {
			frm.mode.value = "addemsprice";
			frm.submit();
		}
	}
}

function CheckBox(frm) {
	if (frm.workidx.value*0 != 0) {
		alert("숫자만 가능합니다.");
		frm.workidx.focus();
		return false;
	}

	return true;
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>Index</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fidx %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>매장ID</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fshopid %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>구분</td>
		<td bgcolor="#FFFFFF" ><font color="<%= ofranchulgomaster.FOneItem.GetDivCodeColor %>"><%= ofranchulgomaster.FOneItem.GetDivCodeName %></font></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>Title</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ftitle %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>총판매가</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>총매입가</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalbuycash,0) %>
		<font color="#AAAAAA">(업체로부터 공급받은 상품가격)</font></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>총공급가</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsuplycash,0) %>
		<font color="#AAAAAA">(샾으로 공급한 상품가격)</font></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>총발행금액</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(ofranchulgomaster.FOneItem.Ftotalsum,0) %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>계산서발행일</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ftaxdate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>입금일</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fipkumdate %></td>
	</tr>
	<% if (CStr(ofranchulgomaster.FOneItem.Fshopdiv) = "7") or (CStr(ofranchulgomaster.FOneItem.Fshopdiv) = "8") then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>상태</td>
		<td bgcolor="#FFFFFF">
			<font color="<%= ofranchulgomaster.FOneItem.GetStateColor %>"><%= ofranchulgomaster.FOneItem.GetStateName %></font>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>관련작업(해외)</td>
		<form name="frmMaster" method="post" action="franmeaippopsubmaster_process.asp">
		<input type="hidden" name="mode" value="<%= mode %>">
		<input type="hidden" name="masteridx" value="<%= idx %>">
		<input type="hidden" name="orgworkidx" value="<%= ofranchulgomaster.FOneItem.Fworkidx %>">
		<td bgcolor="#FFFFFF" >
			<input type="text" class="text" name="workidx" value="<%= ofranchulgomaster.FOneItem.Fworkidx %>" size=6 maxlength=6>
			<input type="button" class="button" value="입력" onClick="UpdateWorkidx(frmMaster)">
			<% if (ofranchulgomaster.FOneItem.Fworkidx <> "") then %>
				<input type="button" class="button" value="조회" onClick="PopExportSheet(<%= ofranchulgomaster.FOneItem.Fworkidx %>)">
				<% if (oCCartoonBoxMasterItem.GetDeliverMethodName = "EMS") then %>
					<font color=red>[운송 : <%= oCCartoonBoxMasterItem.GetDeliverMethodName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		</form>
	</tr>
	<% else %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>상태</td>
		<td bgcolor="#FFFFFF" colspan=3>
			<font color="<%= ofranchulgomaster.FOneItem.GetStateColor %>"><%= ofranchulgomaster.FOneItem.GetStateName %></font>
		</td>
	</tr>
	<% end if %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>기타사항</td>
		<td bgcolor="#FFFFFF" colspan=3>
		<%= nl2Br(ofranchulgomaster.FOneItem.Fetcstr) %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>최초등록자</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Fregusername %>(<%= ofranchulgomaster.FOneItem.Freguserid %>)</td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>최종처리자</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgomaster.FOneItem.Ffinishusername %>(<%= ofranchulgomaster.FOneItem.Ffinishuserid %>)</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="기타내역추가" onClick="popEtcAdd('<%= idx %>','<%= ofranchulgomaster.FOneItem.Fshopid %>')">
			<input type="button" class="button" value="EMS운송비용추가" onClick="popAddDeliverPay(document.frmMaster)" <% if IsNull(oCCartoonBoxMasterItem.GetDeliverMethodName) or (oCCartoonBoxMasterItem.GetDeliverMethodName <> "EMS") then %>disabled<% end if %>>
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<p>

<% if ofranchulgomaster.FOneItem.FDivcode="MC" then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" align=center>
		<td >출고처</td>
		<td width=70>출고코드</td>
		<td width=70>주문코드</td>
		<td width=70>발주일</td>
		<td width=70>출고일</td>
		<td width=80>총판매가</td>
		<td width=80>총매입가</td>
		<td width=80>총공급가</td>
		<td>비고</td>
	</tr>
	<% for i=0 to  ofranchulgojungsan.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + ofranchulgojungsan.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + ofranchulgojungsan.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td ><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= ofranchulgojungsan.FItemList(i).Fshopid %></a></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode01 %></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode02 %></td>
		<td ><%= Left(ofranchulgojungsan.FItemList(i).Fbaljudate, 10) %></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fexecdate %></td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=right><%= FormatNumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
		<td align=center><input type="button" class="button" value="수정" onClick="popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>')"></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td align=right><%= FormatNumber(totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(totalbuycash,0) %></td>
		<td align=right><%= FormatNumber(totalsuplycash,0) %></td>
		<td></td>
	</tr>
</table>

<% elseif ofranchulgomaster.FOneItem.FDivcode="WS" or ofranchulgomaster.FOneItem.FDivcode="TC" or ofranchulgomaster.FOneItem.FDivcode="CC" or ofranchulgomaster.FOneItem.FDivcode="CM" then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" align=center>
		<td>오프샵ID</td>
		<td width=90>판매월</td>
		<td>브랜드ID</td>
		<td width=90>총소비가</td>
		<td width=90>총판매가</td>
		<td width=90>총매입가</td>
		<td width=90>총공급가</td>
		<td>비고</td>
	</tr>
	<% for i=0 to  ofranchulgojungsan.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + ofranchulgojungsan.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + ofranchulgojungsan.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + ofranchulgojungsan.FItemList(i).Ftotalsuplycash
	totalorgsellcash =  totalorgsellcash + ofranchulgojungsan.FItemList(i).Ftotalorgsellcash
	%>
	<tr bgcolor="#FFFFFF">
		<td align=center ><%= ofranchulgojungsan.FItemList(i).Fshopid %></td>
		<td align=center ><a href="javascript:popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>');"><%= ofranchulgojungsan.FItemList(i).Fcode01 %></a></td>
		<td ><%= ofranchulgojungsan.FItemList(i).Fcode02 %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalorgsellcash,0) %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=right><%= Formatnumber(ofranchulgojungsan.FItemList(i).Ftotalsuplycash,0) %></td>
		<td align=center><input type="button" class="button" value="수정" onClick="popSubdetailEdit('<%= ofranchulgojungsan.FItemList(i).Fidx %>','<%= idx %>')"></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td align=right><%= FormatNumber(totalorgsellcash,0) %></td>
		<td align=right><%= FormatNumber(totalsellcash,0) %></td>
		<td align=right><%= FormatNumber(totalbuycash,0) %></td>
		<td align=right><%= FormatNumber(totalsuplycash,0) %></td>
		<td></td>
	</tr>
</table>

<% end if %>
<%
set ofranchulgomaster = Nothing
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->