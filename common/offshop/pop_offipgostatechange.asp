<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 개별 입출고 리스트 상태변경
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim idx ,oipchulmaster
	idx = requestCheckVar(request("idx"),10)

set oipchulmaster = new CShopIpChul
	oipchulmaster.FRectIdx = idx
	oipchulmaster.GetIpChulMasterList
%>

<script type='text/javascript'>
	
function ModiMaster(frm){

	if (frm.statecd[3].checked){
		if (!calendarOpen4(frm.execdate,'입고일',frm.execdate.value)) return;
		var ret = confirm('입고일 : ' + frm.execdate.value + '\n입고 확인 하시겠습니까?');
		if (ret) {
			frm.submit();
			return;
		}
	}

	if (confirm('수정 하시겠습니까?')){
		frm.submit();
	}
}

</script>

<table width="100%" cellspacing="1" cellpadding="2" class="a" bgcolor=#3d3d3d>
<form name="frmMaster" method="post" action="/common/offshop/shopipchul_process.asp">
<input type="hidden" name="mode" value="modistate">
<input type="hidden" name="execdate" value="<%= oipchulmaster.FItemList(0).FexecDt %>">
<input type="hidden" name="idx" value="<%= idx %>">
<tr>
	<td width="100" bgcolor="#DDDDFF">공급처</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="chargeid" value="<%= oipchulmaster.FItemList(0).FChargeid %>">
		<%= oipchulmaster.FItemList(0).FChargeid %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">가맹점 </td>
	<td bgcolor="#FFFFFF">
	<input type="hidden" name="shopid" value="<%= oipchulmaster.FItemList(0).FShopid %>">
		<%= oipchulmaster.FItemList(0).FShopid %> (<%= oipchulmaster.FItemList(0).FShopname %>)
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">총판매가</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSellCash,0) %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">총공급가</td>
	<td bgcolor="#FFFFFF"><%= FormatNumber(oipchulmaster.FItemList(0).FTotalSuplyCash,0) %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">입고예정일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FScheduleDt %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">입고일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).FexecDt %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">가맹점확인일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fshopconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">업체확인일</td>
	<td bgcolor="#FFFFFF">
	<%= oipchulmaster.FItemList(0).Fupcheconfirmdate %>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">등록일</td>
	<td bgcolor="#FFFFFF"><%= oipchulmaster.FItemList(0).FRegDate %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">입고상태</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="statecd" value="-2" <% if oipchulmaster.FItemList(0).Fstatecd="-2" then response.write "checked" %> >입고요청
	<input type="radio" name="statecd" value="-1" <% if oipchulmaster.FItemList(0).Fstatecd="-1" then response.write "checked" %> >입고요청확인 
	<input type="radio" name="statecd" value="0" <% if oipchulmaster.FItemList(0).Fstatecd="0" then response.write "checked" %> >입고대기
	<input type="radio" name="statecd" value="7" <% if oipchulmaster.FItemList(0).Fstatecd="7" then response.write "checked" %> >매장 입고확인
	<input type="radio" name="statecd" value="8" <% if oipchulmaster.FItemList(0).Fstatecd="8" then response.write "checked" %> <% if oipchulmaster.FItemList(0).Fstatecd="0" then response.write "disabled" %> >입고확정
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="button" value="입고상태 수정" onClick="ModiMaster(frmMaster)"></td>
</tr>
</form>
</table>

<%
set oipchulmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->