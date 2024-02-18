<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
rw "관리자문의 요망"
response.end

dim yyyy1,mm1,designer,rectorder, page
designer = request("designer")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
rectorder = request("rectorder")
page = request("page")

if page="" then page=1

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
ojungsan.FCurrPage=page
ojungsan.FPageSize=50
ojungsan.JungsanMasterListSimple

dim i, j
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

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
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

<table width="800" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#FFFFFF">
    	<td colspan=15 align=right>
    	<% for j=1 to ojungsan.FtotalPage %>
	    	<% if page=j then %>
	    	<b><%= j %></b> |
	    	<% else %>
	    	<a href="?page=<%= j %>&rectorder=<%= rectorder %>&designer=<%= designer %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><%= j %></a> |
	    	<% end if %>
    	<% next %>
    	</td>
    </tr>
    <tr bgcolor="#DDDDFF">
      <td width="80"><a href="javascript:research(frm,'designer')">디자이너</a></td>
      <td width="110">Title</td>
      <td width="60">업체배송</td>
      <td width="30">마진</td>
      <td width="60">매입총액</td>
      <td width="30">마진</td>
      <td width="60">특정총액</td>
      <td width="30">마진</td>
      <td width="60">기타판매</td>
      <td width="30">마진</td>
      <td width="60">오프샾</td>
      <td width="30">마진</td>
      <td width="76">총정산액</td>
      <td width="100"><a href="javascript:research(frm,'state')">상태</a></td>
      <td width="30">삭제</td>
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
      <td align="right"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=maeip"><%= FormatNumber(ojungsan.FItemList(i).Fme_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fme_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fme_totalsuplycash/ojungsan.FItemList(i).Fme_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>
      <td align="right"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witaksell"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fwi_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fwi_totalsuplycash/ojungsan.FItemList(i).Fwi_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>

      <td align="right"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witakchulgo"><%= FormatNumber(ojungsan.FItemList(i).Fet_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fet_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fet_totalsuplycash/ojungsan.FItemList(i).Fet_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>

      <td align="right"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(i).FId %>&gubun=witakoffshop"><%= FormatNumber(ojungsan.FItemList(i).Fsh_totalsuplycash,0) %></a></td>
      <% if ojungsan.FItemList(i).Fsh_totalsellcash<>0 then %>
      <td align="right"><%= CLng((1-ojungsan.FItemList(i).Fsh_totalsuplycash/ojungsan.FItemList(i).Fsh_totalsellcash)*10000)/100 %></td>
      <% else %>
      <td align="right">0</td>
      <% end if %>

      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
      <td ><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font>
	  <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      <input type="button" value="진행" onclick="NextStep('<%= ojungsan.FItemList(i).FId %>');">
      <% end if %>
      </td>
      <td align="center">
      <% if ojungsan.FItemList(i).Ffinishflag="0" then %>
      	<a href="javascript:dellThis('<%= ojungsan.FItemList(i).FId %>')">x</a>
      <% end if %>
      </td>
    </tr>
    <% next %>
    <% totsum = totsum + tot1 + tot2 + tot3 + tot4 + tot5 %>
    <tr bgcolor="#FFFFFF">
      <td >합계</td>
      <td ></td>
      <td align="right"><%= FormatNumber(tot1,0) %></td>
      <td ></td>
      <td align="right"><%= FormatNumber(tot2,0) %></td>
      <td ></td>
      <td align="right"><%= FormatNumber(tot3,0) %></td>
      <td ></td>
      <td align="right"><%= FormatNumber(tot4,0) %></td>
      <td ></td>
      <td align="right"><%= FormatNumber(tot5,0) %></td>
      <td ></td>
      <td align="right"><%= FormatNumber(totsum,0) %></td>
      <td ></td>
      <td ></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td colspan=15 align=right>
    	<% for j=1 to ojungsan.FtotalPage %>
	    	<% if page=j then %>
	    	<b><%= j %></b> |
	    	<% else %>
	    	<a href="?page=<%= j %>&rectorder=<%= rectorder %>&designer=<%= designer %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>"><%= j %></a> |
	    	<% end if %>
    	<% next %>
    	</td>
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