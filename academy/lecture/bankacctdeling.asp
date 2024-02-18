<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/lec_bankacctcls.asp"-->

<%
dim ojumun, page, daydiff

daydiff = RequestCheckvar(request("daydiff"),2)
page = RequestCheckvar(request("page"),10)
if page="" then page=1
if daydiff="" then daydiff=10

set ojumun = new CBankAcct
ojumun.FCurrPage = page
ojumun.FPageSize = 100
ojumun.FRectDayDiff = daydiff
ojumun.GetOldMiipkumList

dim i
%>
<script language='javascript'>
function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function delitems(upfrm){
	if (!CheckSelected()){
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택한 무통장 주문을 취소 하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderidx.value = upfrm.orderidx.value + frm.orderidx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" align="right">
			<select name="daydiff">
				<option value="7" <% if daydiff="7" then response.write "selected" %> >7일 이후</option>
				<option value="10" <% if daydiff="10" then response.write "selected" %> >10일 이후</option>
				<option value="15" <% if daydiff="15" then response.write "selected" %> >15일 이후</option>
				<option value="30" <% if daydiff="30" then response.write "selected" %> >30일 이후</option>
			</select>
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<form name="frmarr" method="post" action="/academy/lecture/lib/dobankacct.asp">
<input type="hidden" name="orderidx" value="">
<input type="hidden" name="mode" value="">
<tr>
	<td colspan="13">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
			<tr>
				<td><input type="button" value="선택 주문 취소 처리" onClick="delitems(frmarr)"></td>
				<td align="right">총 건수 : <Font color="#3333FF"><%= FormatNumber(ojumun.FTotalCount,0) %></font></td>
			</tr>
		</table>
	</td>
</tr>
</form>
<tr>
	<td colspan="13" align="right">page : <%= ojumun.FCurrPage %>/<%=ojumun.FTotalPage %></td>
</tr>
<tr >
	<td width="30" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="100" align="center">주문번호</td>
	<td width="80" align="center">Site</td>
	<td width="80" align="center">UserID</td>
	<td width="65" align="center">구매자</td>
	<td width="65" align="center">수령인</td>
	<td width="72" align="center">결제할금액</td>
	<td width="72" align="center">사용마일리지</td>
	<td width="72" align="center">사용쿠폰</td>
	<td width="120" align="center">주문일</td>
	<td width="120" align="center">SMS발송일</td>
</tr>
<% if ojumun.FresultCount<1 then %>
<tr>
	<td colspan="13" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
	<% for i=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" >
	<input type="hidden" name="orderidx" value="<%= ojumun.FItemList(i).FIdx %>">
	<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(i).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td align="center"><a target="_blank" href="/academy/lecture/lec_orderdetail.asp?orderserial=<%= ojumun.FItemList(i).FOrderSerial %>" class="zzz"><%= ojumun.FItemList(i).FOrderSerial %></a></td>
		<td align="center"><%= ojumun.FItemList(i).FSitename %></td>
		<td align="center"><%= ojumun.FItemList(i).FUserID %></td>
		<td align="center"><%= ojumun.FItemList(i).FBuyName %></td>
		<td align="center"><%= ojumun.FItemList(i).FReqName %></td>
		<td align="center"><%= ojumun.FItemList(i).FSubTotalPrice %></td>
		<td align="center"><%= ojumun.FItemList(i).FMileTotalPrice %></td>
		<td align="center"><%= ojumun.FItemList(i).FTenCardSpend %></td>
		<td align="center"><%= ojumun.FItemList(i).FRegDate %>&nbsp;</td>
		<td align="center"><%= ojumun.FItemList(i).FSendRegDate %>&nbsp;</td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="13" height="30" align="center">
		<% if ojumun.HasPreScroll then %>
			<a href="?page=<%= ojumun.StarScrollPage-1 %>&daydiff=<%= daydiff %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ojumun.StarScrollPage to ojumun.FScrollCount + ojumun.StarScrollPage - 1 %>
			<% if i>ojumun.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&daydiff=<%= daydiff %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ojumun.HasNextScroll then %>
			<a href="?page=<%= i %>&daydiff=<%= daydiff %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
<% end if %>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->