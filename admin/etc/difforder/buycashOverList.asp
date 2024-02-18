<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, research, i, page, itemid, makerid, nowsDate, iSD
research	= requestCheckvar(request("research"),2)
itemid  	= request("itemid")
makerid		= requestCheckvar(request("makerid"),32)
iSD			= request("iSD")
'텐바이텐 상품코드 엔터키로 검색되게
If itemid<>"" then
	Dim iA, arrTemp, arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))
	iA = 0
	Do While iA <= ubound(arrTemp)
		If Trim(arrTemp(iA))<>"" then
			If Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			Else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			End If
		End If
		iA = iA + 1
	Loop
	itemid = left(arrItemid,len(arrItemid)-1)
End If
If iSD = "" Then
	nowsDate = Left(dateadd("d",-4,Now()), 7) & "-01"
Else
	nowsDate = Left(iSD, 7) & "-01"
End If

SET oOrder = new COrder
	oOrder.FRectItemid		= itemid
	oOrder.FRectMakerid		= makerid
	oOrder.FRectNowsDate	= nowsDate
	oOrder.getBuycashOverList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function pop_couponView(v){
    var pwin = window.open('/admin/shopmaster/itemcouponlist.asp?menupos=786&research=on&page=1&iSerachType=1&sSearchTxt='+v,'popOutMallEtcLink','width=1200,height=700,scrollbars=yes,resizable=yes');
    pwin.focus();
}
function pop_extsitejungsan(vItemid, vItemCost, vItemoption){
	var pCM5;
	pCM5 = window.open("/admin/etc/extsitejungsan_check.asp?itemid="+vItemid+"&mallsellcash="+vItemCost,"pop_jungsan");	
	pCM5.location.href="/admin/etc/extsitejungsan_check.asp?itemid="+vItemid+"&mallsellcash="+vItemCost+"&itemoption="+vItemoption;
	pCM5.focus();
}
function HighlightRow(obj){
	var table = document.getElementById("tableId");
	var tr = table.getElementsByTagName("tr");
	for(var i=0; i < tr.length; i++){
		tr[i].style.background = "#FFFFFF";
	}
	document.getElementById("topTr").style.background = "#e6e6e6";
	obj.parentElement.style.background = "#FCE6E0";
}
</script>
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<!-- #include virtual="/admin/etc/difforder/gubunTab.asp"-->
<input type="hidden" name="vTab" value="<%= vTab %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2"><h2>※원매입가보다 상품쿠폰 적용 매입가가 큰경우</h2></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		기준월일 : 
		<input id="iSD" name="iSD" value="<%=nowsDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="iSD_trigger" border="0" style="cursor:pointer" align="absbottom" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "iSD", trigger    : "iSD_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		&nbsp;&nbsp;
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		&nbsp;&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
	</td>
	<td align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>
<!-- 리스트 시작 -->
<table id="tableId" width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= FormatNumber(oOrder.FResultCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="topTr">
	<td>사이트명</td>
	<td>상품코드</td>
	<td>옵션코드</td>
	<td>주문번호</td>
	<td>브랜드ID</td>
	<td>판매가</td>
	<td>매입가</td>
	<td>쿠폰적용매입가</td>
	<td>마진체크</td>
	<td>쿠폰번호</td>
	<td>출고일</td>
	<td>정산일</td>
	<td>취소여부</td>
	<td>매입구분</td>
	<td>플러스세일할인</td>
</tr>
<% If oOrder.FResultCount > 0 Then %>
<% For i=0 to oOrder.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= oOrder.FItemList(i).FSitename %></a></td>
	<td><%= oOrder.FItemList(i).FItemid %></td>
	<td><%= oOrder.FItemList(i).FItemoption %></td>
	<td><%= oOrder.FItemList(i).FOrderserial %></td>
	<td><%= oOrder.FItemList(i).FMakerid %></td>
	<td style="cursor:pointer;" onclick="HighlightRow(this);pop_extsitejungsan('<%=oOrder.FItemList(i).FItemID%>', '<%= oOrder.FItemList(i).FItemcost %>', '<%= oOrder.FItemList(i).FItemoption %>' );"><%= oOrder.FItemList(i).FItemcost %></td>
	<td><%= oOrder.FItemList(i).FBuycash %></td>
	<td><%= oOrder.FItemList(i).FbuycashcouponNotApplied %></td>
	<td><%= oOrder.FItemList(i).FChkMargin %></td>
	<td><span style="cursor:pointer;" onclick="pop_couponView('<%=oOrder.FItemList(i).FItemcouponidx%>');"><%= oOrder.FItemList(i).FItemcouponidx %></span></td>
	<td><%= oOrder.FItemList(i).FBeasongdate %></td>
	<td><%= oOrder.FItemList(i).FJungsanFixDate %></td>
	<td><%= oOrder.FItemList(i).FCancelyn %></td>
	<td><%= oOrder.FItemList(i).FOmwdiv %></td>
	<td><%= oOrder.FItemList(i).Fplussalediscount %></td>
</tr>
<% Next %>
<% Else %>
<tr height="50">
    <td colspan="20" align="center" bgcolor="#FFFFFF">
		데이터가 없습니다
    </td>
</tr>
<% End If %>
</table>
<% SET oOrder = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->