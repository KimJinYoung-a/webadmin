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
ojumun.FPageSize = 30
ojumun.FRectDayDiffStart =5
ojumun.FRectDayDiff = daydiff
ojumun.GetMiipkummailingList

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

	var ret = confirm('선택한 무통장 주문 SMS를 발송하시겠습니까?');

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
		upfrm.mode.value="mail";
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
			5일이후~
			<select name="daydiff">
				<option value="10" <% if daydiff="10" then response.write "selected" %> >10일 이전</option>
				<option value="15" <% if daydiff="15" then response.write "selected" %> >15일 이전</option>
			</select>
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
오전 10시경 발송요망, 토일은 제외
<br>
<br>
메세지 내용 : <font color="#CC3333">무통장입금유효기간이틀남았습니다.입금미확인건은자동취소됩니다.더핑거스아카데미^^</font>
<br>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<form name="frmarr" method="post" action="/academy/lecture/lib/dobankacct.asp">
<input type="hidden" name="orderidx" value="">
<input type="hidden" name="mode" value="">
<tr>
	<td colspan="13">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
			<tr>
				<td><input type="button" value="선택주문 SMS발송" onClick="delitems(frmarr)"></td>
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
	<td width="72" align="center">결제할금액</td>
	<td width="72" align="center">사용마일리지</td>
	<td width="72" align="center">핸드폰</td>
	<td width="120" align="center">주문일</td>
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
		<td align="center"><a href="/academy/lecture/lec_orderdetail.asp?orderserial=<%= ojumun.FItemList(i).FOrderSerial %>" class="zzz" target="_blank"><%= ojumun.FItemList(i).FOrderSerial %></a></td>
		<td align="center"><%= ojumun.FItemList(i).FSitename %></td>
		<td align="center"><%= ojumun.FItemList(i).FUserID %></td>
		<td align="center"><%= ojumun.FItemList(i).FBuyName %></td>
		<td align="center"><%= ojumun.FItemList(i).FSubTotalPrice %></td>
		<td align="center"><%= ojumun.FItemList(i).FMileTotalPrice %></td>
		<td align="center"><%= ojumun.FItemList(i).FbuyHp %></td>
		<td align="center"><%= Left(ojumun.FItemList(i).FRegDate,10) %></td>
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