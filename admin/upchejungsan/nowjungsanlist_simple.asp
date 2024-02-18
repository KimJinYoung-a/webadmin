<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim yyyy1,mm1,designer,rectorder
designer = request("designer")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
rectorder = request("rectorder")
dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1
ojungsan.FRectDesigner = designer
ojungsan.FrectOrder = rectorder
ojungsan.JungsanMasterList

dim i
dim tot1,tot2,tot3,tot4,tot5,totsum
tot1 = 0
tot2 = 0
tot3 = 0
tot4 = 0
tot5 = 0
totsum = 0
%>
<script language='javascript'>
function research(frm,order){
	frm.rectorder.value = order;
	frm.submit();
}

function popAcountinfo(v){
	window.open("/admin/lib/popupcheinfo.asp?designer=" + v,"popupcheinfo","width=640 height=500");
}

function popDetail(v){
	window.open('popdetail.asp?id=' + v );
}

function dellThis(v){
	var upfrm = document.frmarr;
	var ret = confirm('모든 정산 데이터를 삭제 하시겠습니까?');
	if (ret){
		upfrm.idx.value = v;
		upfrm.mode.value = "dellall";
		upfrm.submit();
	}
}

function NextStep(idx){
	var upfrm = document.frmarr;
	upfrm.mode.value= "statechange";
	upfrm.idx.value= idx;
	upfrm.rd_state.value="1";

	var ret = confirm('확인 대기 상태로 진행 하시겠습니까?');
	if (ret){
		upfrm.submit();
	}
}
</script>

<table width="1200" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		정산대상년월:<% DrawYMBox yyyy1,mm1 %>
		&nbsp;&nbsp;
		업체:<% drawSelectBoxDesignerwithName "designer",designer  %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="1200" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80"><a href="javascript:research(frm,'designer')">디자이너</a></td>
      <td width="110">Title</td>
      <td width="76">총정산액</td>
      <td width="100"><a href="javascript:research(frm,'state')">상태</a></td>
      <td width="80"><a href="javascript:research(frm,'segum')">세금발행일</a></td>
      <td width="80">입금일</td>
      <td width="50">정산일</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	tot1 = tot1 + ojungsan.FItemList(i).Fub_totalsuplycash
    	tot2 = tot2 + ojungsan.FItemList(i).Fme_totalsuplycash
    	tot3 = tot3 + ojungsan.FItemList(i).Fwi_totalsuplycash
    	tot4 = tot4 + ojungsan.FItemList(i).Fet_totalsuplycash
    	tot5 = tot5 + ojungsan.FItemList(i).Fsh_totalsuplycash
    %>
    <tr bgcolor="#FFFFFF">
      <td ><a href="javascript:popAcountinfo('<%= ojungsan.FItemList(i).Fdesignerid %>')"><%= ojungsan.FItemList(i).Fdesignerid %></a></td>
      <td ><a href="nowjungsanmasteredit.asp?id=<%= ojungsan.FItemList(i).FId %>"><%= ojungsan.FItemList(i).Ftitle %></a></td>
      <td align="right"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=upche"><%= FormatNumber(ojungsan.FItemList(i).Fub_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fub_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fub_totalsuplycash/ojungsan.FItemList(i).Fub_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font>
      </td>
      <% if isNull(ojungsan.FItemList(i).Ftaxregdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Ftaxregdate),10) %></td>
      <% end if %>
      <% if isNull(ojungsan.FItemList(i).Fipkumdate) then %>
      <td ></td>
      <% else %>
      <td ><%= Left(Cstr(ojungsan.FItemList(i).Fipkumdate),10) %></td>
      <% end if %>
      <td align="center"><%= ojungsan.FItemList(i).Fjungsan_date %></td>
    </tr>
    <% next %>
    <% totsum = totsum + tot1 + tot2 + tot3 + tot4 + tot5 %>
    <tr bgcolor="#FFFFFF">
      <td >합계</td>
      <td ></td>
      <td align="right"><%= FormatNumber(totsum,0) %></td>
      <td ></td>
      <td ></td>
      <td ></td>
      <td ></td>
    </tr>
</table>
<form name="frmarr" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="rd_state" value="">
</form>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->