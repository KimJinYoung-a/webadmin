<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/difforder/diffOrderCls.asp"-->
<%
Dim oOrder, research, i, page, itemid, makerid, nowsDate,iSD
research	= requestCheckvar(request("research"),2)
'itemid  	= request("itemid")
'makerid		= requestCheckvar(request("makerid"),32)
iSD			= request("iSD")
'텐바이텐 상품코드 엔터키로 검색되게
' If itemid<>"" then
' 	Dim iA, arrTemp, arrItemid
' 	itemid = replace(itemid,",",chr(10))
' 	itemid = replace(itemid,chr(13),"")
' 	arrTemp = Split(itemid,chr(10))
' 	iA = 0
' 	Do While iA <= ubound(arrTemp)
' 		If Trim(arrTemp(iA))<>"" then
' 			If Not(isNumeric(trim(arrTemp(iA)))) then
' 				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
' 				dbget.close()	:	response.End
' 			Else
' 				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
' 			End If
' 		End If
' 		iA = iA + 1
' 	Loop
' 	itemid = left(arrItemid,len(arrItemid)-1)
' End If
If iSD = "" Then
	nowsDate = Left(dateadd("d",-4,Now()), 7) & "-01"
Else
	nowsDate = Left(iSD, 7) & "-01"
End If

SET oOrder = new COrder
	'oOrder.FRectItemid		= itemid
	'oOrder.FRectMakerid		= makerid
	oOrder.FRectNowsDate	= nowsDate
	oOrder.getTaxErrList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function pop_optionEdit(v){
    var pwin = window.open('/common/pop_simpleitemedit.asp?itemid='+v,'popOutMallEtcLink','width=500,height=700,scrollbars=yes,resizable=yes');
    pwin.focus();
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
	<td colspan="2">
		<h2>
		※브랜드는 면세가 아닌데, 상품이 면세로 되어있는 경우<br />
		1. 상품 과세 => 면세 수정<br />
		2. 주문 내역 수정
		</h2>
	</td>
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
		<% if (FALSE) then %>
		&nbsp;&nbsp;
		상품코드 : <textarea rows="2" cols="20" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		
		&nbsp;&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
		<% end if %>
	</td>
	<td align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</table>
</form>

<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		검색결과 : <b><%= FormatNumber(oOrder.FResultCount,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>상품코드</td>
	<td>브랜드ID</td>
	<td>대카테코드</td>
	<td>중카테코드</td>
	<td>소카테코드</td>
	<td>대카테명</td>
	<td>중카테명</td>
	<td>소카테명</td>
</tr>
<% If oOrder.FResultCount > 0 Then %>
<% For i=0 to oOrder.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
	<% if oOrder.FItemList(i).FItemid<>-1 then %>
	<a href="/admin/itemmaster/itemlist.asp?itemid=<%= oOrder.FItemList(i).FItemid %>" target="_popitemlist"><%= oOrder.FItemList(i).FItemid %></a>
	<% end if %>
	</td>
	<td>
	<% if oOrder.FItemList(i).FMakerid<>"" then %>
	<a href="/admin/itemmaster/itemlist.asp?makerid=<%= oOrder.FItemList(i).FMakerid %>" target="_popitemlist"><%= oOrder.FItemList(i).FMakerid %></a>
	<% end if %>
	</td>
	<td><%= oOrder.FItemList(i).FCate_large %></td>
	<td><%= oOrder.FItemList(i).FCate_mid %></td>
	<td><%= oOrder.FItemList(i).FCate_small %></td>
	<td><%= oOrder.FItemList(i).FNmlarge %></td>
	<td><%= oOrder.FItemList(i).FNmmid %></td>
	<td><%= oOrder.FItemList(i).FNmsmall %></td>
</tr>
<% Next %>
<% Else %>
<tr height="50">
    <td colspan="18" align="center" bgcolor="#FFFFFF">
		데이터가 없습니다
    </td>
</tr>
<% End If %>
</table>
<% SET oOrder = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->