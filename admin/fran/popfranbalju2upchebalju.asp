<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->

<%
dim designerid, baljuid
designerid  = RequestCheckVar(request("designerid"),32)
baljuid     = RequestCheckVar(request("baljuid"),32)

dim oordersheet2


set oordersheet2 = new COrderSheet
oordersheet2.FPageSize = 50
oordersheet2.FRectBaljuId = baljuid
oordersheet2.FRectMakerid = designerid
'oordersheet2.FRectStatecd = "1"

if designerid<>"" then
	oordersheet2.GetFranBalju2UpcheBaljuSheetList
end if

dim i
%>
<script language='javascript'>
function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popshopjumunsheet.asp?idx=' + v + '&itype=' + itype,'shopjumunsheet','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function MakeJumun(idesignerid){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	var idxarr="";
	var etcstr="";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				idxarr = idxarr + frm.idx.value + ",";
				etcstr = etcstr + frm.etcstr.value + ",";
			}
		}
	}

	if (confirm('주문서를 작성하시겠습니까?')){
		opener.MakeJumunByIdx(idxarr,idesignerid,etcstr);
		close();
	}
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="baljuid" value="<%= baljuid %>" >

	<tr>
		<td class="a" >
			브랜드 :
			<% drawSelectBoxDesignerwithName "designerid",designerid %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr height=30>
	<td><b>최근 <%= oordersheet2.FPageSize %>건 주문 역순 정렬</b></td>
	<td align=right><input type=button value="선택주문으로 주문서 작성" onclick="MakeJumun('<%= designerid %>');"></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td width=20><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
	<td width=60>주문코드</td>
	<td width=64>주문일</td>
	<td width=64>입고요청일</td>
	<td width=120>발주처</td>
	<td width=60>소비가</td>
	<td width=60>공급가</td>
	<td>구분</td>
	<td>상태</td>
</tr>
<% for i=0 to oordersheet2.FResultCount -1 %>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="idx" value="<%= oordersheet2.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><a href="javascript:PopIpgoSheet('<%= oordersheet2.FItemList(i).Fidx %>','3')"><%= oordersheet2.FItemList(i).FBaljuCode %></a></td>
	<td ><%= Left(oordersheet2.FItemList(i).FRegdate,10) %></td>
	<td ><%= Left(oordersheet2.FItemList(i).FScheduleDate,10) %></td>
	<td ><%= oordersheet2.FItemList(i).FBaljuid %> [<%= oordersheet2.FItemList(i).FBaljuName %>]</td>
	<td align=right>
	<%= FormatNumber(oordersheet2.FItemList(i).Fjumunsellcash,0) %>
	<br>
	<%= FormatNumber(oordersheet2.FItemList(i).Ftotalsellcash,0) %>
	</td>
	<td align=right>
	<%= FormatNumber(oordersheet2.FItemList(i).Fjumunbuycash,0) %>
	<br>
	<%= FormatNumber(oordersheet2.FItemList(i).Ftotalbuycash,0) %>
	</td>
	<td><%= oordersheet2.FItemList(i).GetDivCodeName %></td>
	<td><font color="<%= oordersheet2.FItemList(i).GetStateColor %>"><%= oordersheet2.FItemList(i).GetStateName %></font></td>
	<input type=hidden name="etcstr" value="<%= oordersheet2.FItemList(i).FBaljuCode %>[<%= oordersheet2.FItemList(i).FBaljuName %>]">
</tr>
</form>
<% next %>
</table>

<%
set oordersheet2 = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->