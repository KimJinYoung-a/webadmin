<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/bankacctcls.asp"-->
<%
dim ojumun, page, daydiff

daydiff = request("daydiff")
page = requestCheckvar(request("page"),10)
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

    <% if not(C_ADMIN_AUTH) then %>
    alert('더이상 지원하지 않는 메뉴 입니다. - 매일 오전 자동으로 발송됨.');
    return;
    <% else %>
    alert('관리자 권한 강제실행');
    <% end if %>
	if (!CheckSelected()){
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택한 무통장 주문 메일을 발송하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderidx.value = upfrm.orderidx.value + frm.orderidx.value + ",";
					upfrm.orderSerialArray.value = upfrm.orderSerialArray.value + frm.orderserial.value + "," ;

				}
			}
		}
		upfrm.mode.value="mail";
		upfrm.submit();

	}
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			5일이후~
			<select class="select" name="daydiff">
				<option value="10" <% if daydiff="10" then response.write "selected" %> >10일 이전</option>
				<option value="15" <% if daydiff="15" then response.write "selected" %> >15일 이전</option>
				<!-- option value="55" <% if daydiff="55" then response.write "selected" %> >55일 이전</option -->
			</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmarr" method="post" action="dobankacct.asp">
	<input type="hidden" name="orderidx" value="">
	<input type="hidden" name="orderSerialArray" value="">
	<input type="hidden" name="mode" value="">
	<tr>
		<td align="left">
			<input type="button" class="button" value="선택주문 메일발송" onClick="delitems(frmarr)">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= FormatNumber(ojumun.FTotalCount,0) %></b>
			&nbsp;
			페이지 : <b<%= ojumun.FCurrPage %> / <%=ojumun.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="100" align="center">주문번호</td>
		<td width="80" align="center">Site</td>
		<td width="80" align="center">UserID</td>
		<td width="65" align="center">구매자</td>
		<td width="65" align="center">수령인</td>
		<td width="72" align="center">결제할금액</td>
		<td width="72" align="center">사용마일리지</td>
		<td width="65" align="center">보조결제총액</td>
		<td width="72" align="center">이메일</td>
		<td width="40" align="center">가상</td>
		<td width="120" align="center">입금은행</td>
		<td width="120" align="center">주문일</td>
	</tr>
<% if ojumun.FresultCount<1 then %>
	<tr>
		<td colspan="15" align="center" bgcolor="FFFFFF">[검색결과가 없습니다.]</td>
	</tr>
<% else %>
	<% for i=0 to ojumun.FresultCount-1 %>
	<form name="frmBuyPrc_<%=i%>" method="post" >
	<input type="hidden" name="orderidx" value="<%= ojumun.FItemList(i).FIdx %>">
	<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(i).FOrderSerial %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td align="center"><a href="#" onclick="ViewOrderDetail(frmBuyPrc_<%=i%>)" class="zzz"><%= ojumun.FItemList(i).FOrderSerial %></a></td>
		<td align="center"><%= ojumun.FItemList(i).FSitename %></td>
		<td align="center"><%= printUserId(ojumun.FItemList(i).FUserID,2,"**") %></td>
		<td align="center"><%= ojumun.FItemList(i).FBuyName %></td>
		<td align="center"><%= ojumun.FItemList(i).FReqName %></td>
		<td align="center"><%= ojumun.FItemList(i).FSubTotalPrice-ojumun.FItemList(i).FSumPaymentEtc %></td>
		<td align="center"><%= ojumun.FItemList(i).FMileTotalPrice %></td>
		<td align="center"><%= ojumun.FItemList(i).FSumPaymentEtc %></td> 
		<td align="center"><%= ojumun.FItemList(i).FbuyEmail %></td>
		<td align="center"><%= CHKIIF(ojumun.FItemList(i).IsDacomCyberAccountPay,"가상","일반") %></td>
		<td align="center"><%= ojumun.FItemList(i).FAccountNo %></td>
		<td align="center"><%= Left(ojumun.FItemList(i).FRegDate,10) %></td>
	</tr>
	</form>
	<% next %>
	<tr>
		<td colspan="15" height="30" align="center">
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