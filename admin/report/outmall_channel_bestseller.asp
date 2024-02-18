<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/category_reportcls.asp"-->
<%
Dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
Dim nowdate,searchnextdate
Dim oreport
Dim topn,depth1,page
Dim ix, cknodate
Dim order_desum
Dim ordertype
Dim oldlist, sitename
Dim totalsumprice, totalbuyprice, totalitemno
Dim GpRdsite

yyyy1		= request("yyyy1")
mm1			= request("mm1")
dd1			= request("dd1")
yyyy2		= request("yyyy2")
mm2			= request("mm2")
dd2			= request("dd2")
depth1		= request("depth1")
topn		= request("topn")
cknodate	= request("cknodate")
order_desum	= request("order_desum")
ordertype	= request("ordertype")
If ordertype = "" Then ordertype = "ea"		'디폴트 수량순 정렬
oldlist		= request("oldlist")
sitename	= request("sitename")
GpRdsite	= request("GpRdsite")
If (yyyy1 = "") Then
	nowdate = Left(CStr(now()),10)
	yyyy1	= Left(nowdate,4)
	mm1 	= Mid(nowdate,6,2)
	dd1 	= Mid(nowdate,9,2)

	yyyy2	= yyyy1
	mm2		= mm1
	dd2		= dd1
Else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1	= Left(nowdate,4)
	mm1		= Mid(nowdate,6,2)
	dd1		= Mid(nowdate,9,2)
end if
searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

topn = request("topn")
If (topn = "") Then topn = 100
Set oreport = new CCategoryReport
If cknodate = "" Then
	oreport.FRectFromDate = yyyy1 & "-" & mm1 & "-" + dd1
	oreport.FRectToDate = searchnextdate
end if
	oreport.FRectDepth1		= depth1
	oreport.FPageSize		= topn
	oreport.FCurrPage		= page
	oreport.FRectOrdertype	= ordertype
	oreport.FRectOldJumun	= oldlist
	oreport.FRectSitename	= sitename
	oreport.FRectGpRdsite	= GpRdsite
	oreport.OutmallSearchCategoryBestseller
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
function ReSearch(ifrm){
	var v = ifrm.topn.value;
	if (!IsDigit(v)){
		alert('숫자만 가능합니다.');
		ifrm.topn.focus();
		return;
	}

	if (v>1000){
		alert('천건 이하만 검색가능합니다.');
		ifrm.topn.focus();
		return;
	}
	ifrm.submit();
}
function jsGrouplist(nm){
	$("#GroupList").empty();
	var str = $.ajax({
		type: "POST",
		url: "/admin/report/act_GroupListajax.asp",
		data: "gubun="+nm+"&GpRdsite=<%=GpRdsite%>",
		dataType: "text",
		async: false
	}).responseText;

	if(str!="") {
		$("#GroupList").html(str);
	}
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr>
	<td class="a" >
	<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
	기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	카테고리선택 : <% DrawSelectBoxDispCateLarge "depth1", depth1, "" %>&nbsp;
	검색갯수 :
	<input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6" >
	사이트구분 : <% Drawsitename "sitename", sitename, "Y" %>&nbsp;
	<span id="GroupList">
	
	</span>
	<br>
	&nbsp;&nbsp;
	<input type="radio" name="ordertype" value="ea" <% if ordertype="ea" then response.write "checked" %>>수량순
	<input type="radio" name="ordertype" value="totalprice" <% if ordertype="totalprice" then response.write "checked" %>>매출순
	<input type="radio" name="ordertype" value="gain" <% if ordertype="gain" then response.write "checked" %>>수익순
	<input type="radio" name="ordertype" value="unitCost" <% if ordertype="unitCost" then response.write "checked" %>>객단가순
	<br>
	</td>
	<td class="a" align="right">
		<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#E6E6E6">
	<td colspan="12" height="25" align="right">검색결과 : 총 <font color="red"><% = oreport.FResultCount %></font>개</td>
</tr>
<tr bgcolor="#E6E6E6">
	<td width="30" align="center">순위</td>
	<td width="50" align="center">이미지</td>
	<td width="50" align="center">상품번호</td>
	<td  align="center">상품</td>
	<td width="50">단가</td>
	<td width="100" align="center">브랜드ID</td>
	<td width="80" align="center">옵션</td>
	<td width="65" align="center">판매갯수</td>
	<td width="100" align="center">판매가합</td>
	<td width="100" align="center">매입가합</td>
	<td width="100" align="center">수익</td>
	<td width="70" align="center">마진율</td>
</tr>
<% If oreport.FResultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center">[검색결과가 없습니다.]</td>
</tr>
<%
   Else
	For ix = 0 to oreport.FResultCount -1
		totalitemno   =  totalitemno + oreport.FItemList(ix).FItemNo
		totalsumprice =  totalsumprice + oreport.FItemList(ix).Fselltotal
		totalbuyprice =  totalbuyprice + oreport.FItemList(ix).Fbuytotal
%>
	<tr class="a" bgcolor="#FFFFFF" height="50">
		<td align="center"><%=ix+1%></td>
		<td><img src="<%= oreport.FItemList(ix).FImageSmall %>" width=50></td>
		<td align="center" height="25"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oreport.FItemList(ix).FItemID %>" class="zzz" target="_blank"><%= oreport.FItemList(ix).FItemID  %></a></td>
		<td align="center"><%= oreport.FItemList(ix).FItemName %></td>
		<td align="center"><%= FormatNumber(oreport.FItemList(ix).FItemCost,0) %></td>
		<td align="center"><%= oreport.FItemList(ix).FMakerid %></td>
		<% if (oreport.FItemList(ix).FItemOptionStr="") then %>
			<td align="center">&nbsp;</td>
		<% else %>
			<td align="center"><%= oreport.FItemList(ix).FItemOptionStr %></td>
		<% end if %>
		<td align="center"><%= oreport.FItemList(ix).FItemNo %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fbuytotal,0) %></td>
		<td align="right"><%= FormatNumber(oreport.FItemList(ix).Fselltotal-oreport.FItemList(ix).Fbuytotal,0) %></td>
	    <td align="center">
	        <% if oreport.FItemList(ix).Fselltotal<>0 then %>
	        <%= 100-CLng(oreport.FItemList(ix).Fbuytotal/oreport.FItemList(ix).Fselltotal*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" align="center">Total</td>
	    <td colspan="5"></td>
	    <td align="center"><%= FormatNumber(totalitemno,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalbuyprice,0) %></td>
	    <td align="right"><%= FormatNumber(totalsumprice-totalbuyprice,0) %></td>
	    <td align="center">
	        <% if totalsumprice<>0 then %>
	        <%= 100-CLng(totalbuyprice/totalsumprice*100*100)/100 %> %
	        <% end if %>
	    </td>
	</tr>
<% end if %>
</table>
<script language="javascript">
	jsGrouplist("<%=sitename%>");
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->