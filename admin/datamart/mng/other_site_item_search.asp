<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include file="./other_site_iteminfo_cls.asp" -->
<!-- #include virtual="/lib/classes/search/itemCls.asp" -->
<!-- #include virtual="/lib/classes/search/searchCls.asp" -->
<%
	Dim i, cOSI, vPage, vSiteCode, vSiteItemID, vSearchTxt, vItemID, vItemName, vRegDate, vSiteItemName
	Dim sortmethod, lp, dispCate, makerid, sellyn
	vPage = NullFillWith(requestCheckVar(request("page"),10),1)
	vSiteCode = requestCheckVar(request("sitecode"),50)
	vSiteItemID = requestCheckVar(request("siteitemid"),15)
	vSearchTxt = requestCheckVar(request("searchtxt"),100)
	vRegDate = requestCheckVar(request("regdate"),10)

	'### ∞Àªˆø£¡¯
	vSearchTxt = RepWord(vSearchTxt,"[^∞°-∆Ra-zA-Z0-9.&%\-\_\s]","")

	If sortmethod = "" Then sortmethod = "ne"

	SET cOSI = new COSItem
	cOSI.FRectSiteCode = vSiteCode
	cOSI.FRectSiteItemID = vSiteItemID
	cOSI.FRectRegDate = vRegDate
	cOSI.sbOtherSiteItem
	vItemID = cOSI.FOneItem.Fitemid
	vItemName = cOSI.FOneItem.Fitemname
	vSiteItemName = cOSI.FOneItem.Fsiteitemname
	SET cOSI = Nothing
	
	If vSearchTxt = "" Then
		vSearchTxt = vSiteItemName
	End If

	Dim oDoc
	Set oDoc = new SearchItemCls
		oDoc.FCurrPage = vPage
		oDoc.FPageSize = 15
		oDoc.FScrollCount = 10
		oDoc.FRectSearchTxt		= vSearchTxt
		oDoc.FRectCateCode		= dispCate				'ƒ´≈◊∞Ì∏Æƒ⁄µÂ
		oDoc.FRectMakerid		= makerid				'æ˜√º æ∆¿Ãµ
		oDoc.FListDiv			= "fulllist"
		oDoc.FSellScope			= sellyn
		oDoc.FRectSortMethod	= sortmethod
		oDoc.getSearchList

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<%
	
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<style type="text/css">
.fontred {color:#FF0000 !important;}
</style>
<script>
function searchFrm(f,p){
	if(f == "frm1"){
		frm1.page.value = p;
	}else{
		frm2.page.value = p;
	}
	jsItemlist();
}
function jsItemlist(){
	$("#frm1").submit();
}
function jsItemSave(i){
	$("#action").val("");
	$("#itemid").val(i);
	$("#procfrm").submit();
}
function jsItemDelete(){
	$("#action").val("delete");
	$("#procfrm").submit();
}
</script>

<div class="contSectFix scrl" id="linetop">
	&nbsp;<strong>ªÁ¿Ã∆Æ : <%=vSiteCode%>, ªÁ¿Ã∆ÆªÛ«∞ƒ⁄µÂ : <%=vSiteItemID%></strong>
	<hr  color="#000000;" />
	<br />
	<form name="frm1" id="frm1" method="get">
	<input type="hidden" name="page" id="page" value="<%=vPage%>">
	<input type="hidden" name="sitecode" id="sitecode" value="<%=vSiteCode%>">
	<input type="hidden" name="siteitemid" id="siteitemid" value="<%=vSiteItemID%>">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li class="lMar10 rMar10">
					<span class="tPad10"><label class="formTit bold" for="term1">ªÛ«∞∏Ì ∞Àªˆ</label></span>
				</li>
				<li class="lMar10 rMar10">
					<input type="text" name="searchtxt" id="searchtxt" size="60" style="height:25px;" value="<%=vSearchTxt%>" onKeyPress="if (event.keyCode == 13){ jsItemlist(); return false;}">
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="∞Àªˆ" onClick="jsItemlist();" />
	</div>
	</form>
	<div class="pad20" id="itemserch">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* √— <%= FormatNumber(oDoc.FTotalCount,0) %> ∞≥</p>
			</div>
			<div class="ftRt">
			<% If vItemID <> "" Then %>
				<span id="matchingspan"><font color="blue"><strong>∏≈ƒ™µ» ªÛ«∞ : [<%=vItemID%>] <%=vItemName%></strong></font></span>
				<input type="button" value="∏≈ƒ™ªË¡¶" onClick="jsItemDelete();">
			<% End If %>
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>ªÛ«∞ƒ⁄µÂ</div></th>
					<th><div>¿ÃπÃ¡ˆ</div></th>
					<th><div>∫Í∑£µÂID</div></th>
					<th><div>ªÛ«∞∏Ì</div></th>
					<th><div></div></th>
				</tr>
				</thead>
				<tbody>
				<% If oDoc.FResultCount > 0 Then %>
					<% For i = 0 To oDoc.FResultCount - 1 %>
					<tr>
						<td>
							[<%= oDoc.FItemList(i).FItemID %>] [<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oDoc.FItemList(i).FItemID %>" target="_blank">∏µ≈©</a>]
						</td>
						<td><img src="<%= oDoc.FItemList(i).FImageSmall %>"></td>
						<td><%= oDoc.FItemList(i).FMakerid %></td>
						<td><%= oDoc.FItemList(i).FItemname %></td>
						<td><input type="button" value="¿˙¿Â" onClick="jsItemSave('<%= oDoc.FItemList(i).FItemID %>');"></td>
					</tr>
					<% Next %>
				<% Else %>
					<tr>
						<td colspan="5">∞Àªˆµ» ªÛ«∞¿Ã æ¯Ω¿¥œ¥Ÿ.</td>
					</tr>
				<% End If %>
				</tbody>
			</table>
		</div>
		<br />
		<div class="ct tPad20 cBk1">
			<% if oDoc.HasPreScroll then %>
			<a href="javascript:searchFrm('frm1','<%= oDoc.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
			
			<% for i=0 + oDoc.StartScrollPage to oDoc.FScrollCount + oDoc.StartScrollPage - 1 %>
				<% if i>oDoc.FTotalpage then Exit for %>
				<% if CStr(vPage)=CStr(i) then %>
				<span class="cRd1">[<%= i %>]</span>
				<% else %>
				<a href="javascript:searchFrm('frm1','<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>
			
			<% if oDoc.HasNextScroll then %>
				<a href="javascript:searchFrm('frm1','<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</div>
	</div>
	
	<!--
	<hr  color="#000000;" />
	<br />
	
	<form name="frm2" method="get" action="">
	<input type="hidden" name="page" value="">
	<div class="searchWrap">
		<div class="search rowSum1">
			<ul>
				<li class="lMar10 rMar10">
					<span class="tPad10"><label class="formTit bold" for="term1">∫Í∑£µÂ ∞Àªˆ</label></span>
				</li>
				<li class="lMar10 rMar10">
				</li>
			</ul>
		</div>
		<input type="submit" class="schBtn" value="∞Àªˆ" />
	</div>
	</form>
	<div class="pad20" style="height:300px;overflow:auto;" id="brandserch">
		<div class="overHidden">
			<div class="ftLt">
				<p class="cBk1 ftLt">* √—  ∞≥</p>
			</div>
			<div class="ftRt">
				
			</div>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<thead>
				<tr>
					<th><div>µÓ∑œ¿œ</div></th>
				</tr>
				</thead>
				<tbody>
				<tr>
					<td></td>
				</tr>
				</tbody>
			</table>
		</div>
		<br />
	</div>
	-->
</div>
<form name="procfrm" id="procfrm" action="other_site_item_search_proc.asp" method="post" style="margin:0px;" target="prociframe">
<input type="hidden" name="action" id="action" value="">
<input type="hidden" name="itemid" id="itemid" value="">
<input type="hidden" name="sitecode" id="sitecode" value="<%=vSiteCode%>">
<input type="hidden" name="siteitemid" id="siteitemid" value="<%=vSiteItemID%>">
</form>
<iframe name="prociframe" id="prociframe" src="about:blank" width="0" height="0" style="margin:0px;"></iframe>
</html>
<% Set oDoc = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->