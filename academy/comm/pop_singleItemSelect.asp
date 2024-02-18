<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%'' #include virtual="/lib/classes/items/itemcls_2008.asp"%>
<!-- #include virtual="/academy/lib/classes/DIYShopitem/DIYitemCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
'###############################################
' PageName : pop_singleItemSelect.asp
' Discription : 상품 단품 선택 팝업
' 	간단히 한가지 상품을 검색해서 선택 반환
'		사용페이지: category_main_pageItem_input.asp
' History : 2008.04.08 허진원 : 생성 
' History : 2016.08.01 유태욱 : 핑거스용 
'###############################################

dim makerid,itemid,itemname, sortDiv
dim page,sellyn,packyn
dim target, ptype
dim cdl,cdm,cds, sailyn
Dim dispCate : dispCate = RequestCheckvar(Request("disp"),10)

cdl = RequestCheckvar(request("cdl"),10)
cdm = RequestCheckvar(request("cdm"),10)
cds = RequestCheckvar(request("cds"),10)
makerid = RequestCheckvar(request("makerid"),32)
itemid = RequestCheckvar(request("itemid"),10)
itemname = RequestCheckvar(request("itemname"),64)
sellyn = RequestCheckvar(request("sellyn"),1)
page = RequestCheckvar(request("page"),10)
target= RequestCheckvar(request("target"),16)
ptype= RequestCheckvar(request("ptype"),16)
sailyn = RequestCheckvar(request("sailyn"),1)
sortDiv = RequestCheckvar(request("sortDiv"),10)

if page="" then page=1 
if sellyn = "" then sellyn = "Y"
if sortDiv = "" then sortDiv = "new"
dim oItem
set oItem = new CItem
oItem.FCurrPage = page
oItem.FPageSize = 20
oItem.FRectItemName = itemname
oItem.FRectMakerid = makerid
oItem.FRectItemid = itemid
oItem.FRectSellYN = sellyn
oItem.FRectCate_Large = cdl
oItem.FRectCate_Mid = cdm
oItem.FRectCate_Small = cds
oItem.FRectDispCate = dispCate
oItem.FRectSailYn = sailyn
oItem.FRectSortDiv = sortDiv
oItem.GetItemList

dim i
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script>
<% If ptype = "just1day" or ptype = "mdpick" Then %>
document.domain = "10x10.co.kr";
<% End If %>

function SelectItems(tgvalue,Nm,oP,sP,sC,Ln,Ly,oC,Bi){
	var tg = eval('opener.<%= target %>');

	tg.itemid.value = tgvalue;
	<%
		Select Case ptype
			Case "just1day"
	%>
	tg.orgPrice.value = oP;
	tg.salePrice.value = sP;
	tg.saleSuplyCash.value = sC;
	tg.limitNo.value = Ln;
	tg.limitYn.value = Ly;
	tg.itemOptCnt.value = oC;
	tg.image1.value = Bi;
	opener.putPercent();
	<%		Case "CateMainPage" %>
	tg.all.itemname.innerText = Nm;
	<% end Select %>
	self.close();
}

function goPage(pg) {
	document.frm.page.value=pg;
	document.frm.submit();
}

function jsSortThis(s){
	document.frm.sortDiv.value=s;
	document.frm.submit();
}
</script>
<body bgcolor="#F4F4F4">
<!-- 해더 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
		<tr>
			<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
				<font color="#333333"><b>상품 선택/추가</b></font>
			</td>
			<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
		상품을 검색하고 한가지 상품을 선택합니다.
	</td>
</tr>
<tr><td height="10"></td></tr>
</table>
<!-- 검색 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="target" value="<%= target %>">
<input type="hidden" name="ptype" value="<%= ptype %>">
<input type="hidden" name="sortDiv" value="<%= sortDiv %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
			디자이너 :
			<% drawSelectBoxDesigner "makerid",makerid %>
			상품ID :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="8" maxlength="32">
			할인 :
			   <select class="select" name="sailyn">
			   <option value="">전체</option>
			   <option value="Y"  <%=CHKIIF(sailyn="Y","selected","")%>>할인</option>
			   <option value="N"  <%=CHKIIF(sailyn="N","selected","")%>>할인안함</option>
			   </select>
			</td>
		</tr>
		<tr>
			<td>
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
				&nbsp;판매여부 :
				<select name="sellyn" class="select">
			                 	<option value='' selected>선택</option>
			                 	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
			                 	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
			         	</select>
			    <br>
			    전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
			</td>
		</tr>
		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 본문 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%= FormatNumber(oItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oItem.FTotalPage,0) %></b>
	</td>
	<td colspan="2">정렬:
		<select name="sort" onchange="jsSortThis(this.value);">
		<option value="new" <%=CHKIIF(sortDiv="new","selected","")%>>신상품순</option>
		<option value="cashH" <%=CHKIIF(sortDiv="cashH","selected","")%>>높은가격순</option>
		<option value="cashL" <%=CHKIIF(sortDiv="cashL","selected","")%>>낮은가격순</option>
		<option value="best" <%=CHKIIF(sortDiv="best","selected","")%>>베스트순</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">상품명</td>
	<td align="center">가격</td>
	<td align="center">업체ID</td>
	<td align="center">배송구분</td>
	<td align="center">판매여부</td>
</tr>
<% for i=0 to oItem.FresultCount-1 %>
<form name="frmBuyPrc_<%= oItem.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="doitemviewset.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= oItem.FItemList(i).FItemID %>">
<tr bgcolor="#FFFFFF" align="center">
	<td><a href="javascript:SelectItems(<%= oItem.FItemList(i).FItemID %>,'<%= Replace(oItem.FItemList(i).FItemName,"'","\'") %>',<%= oItem.FItemList(i).ForgPrice %>,<%= oItem.FItemList(i).FsailPrice %>,<%= oItem.FItemList(i).FsailSuplyCash%>,<%= oItem.FItemList(i).FlimitNo %>,'<%= oItem.FItemList(i).FlimitYn %>',<%=oItem.FItemList(i).FoptionCnt%>,'<%= oItem.FItemList(i).Fbasicimage %>')"><%= oItem.FItemList(i).FItemID %></a></td>
	<td><a href="javascript:SelectItems(<%= oItem.FItemList(i).FItemID %>,'<%= Replace(oItem.FItemList(i).FItemName,"'","\'") %>',<%= oItem.FItemList(i).ForgPrice %>,<%= oItem.FItemList(i).FsailPrice %>,<%= oItem.FItemList(i).FsailSuplyCash%>,<%= oItem.FItemList(i).FlimitNo %>,'<%= oItem.FItemList(i).FlimitYn %>',<%=oItem.FItemList(i).FoptionCnt%>,'<%= oItem.FItemList(i).Fbasicimage %>')"><img src="<%= oItem.FItemList(i).Fsmallimage %>" width="50" height="50" border=0></a></td>
	<td align="left"><a href="javascript:SelectItems(<%= oItem.FItemList(i).FItemID %>,'<%= Replace(oItem.FItemList(i).FItemName,"'","\'") %>',<%= oItem.FItemList(i).ForgPrice %>,<%= oItem.FItemList(i).FsailPrice %>,<%= oItem.FItemList(i).FsailSuplyCash%>,<%= oItem.FItemList(i).FlimitNo %>,'<%= oItem.FItemList(i).FlimitYn %>',<%=oItem.FItemList(i).FoptionCnt%>,'<%= oItem.FItemList(i).Fbasicimage %>')"><%= oItem.FItemList(i).FItemName %></a></td>
	<td align="left"><%= FormatNumber(oItem.FItemList(i).Fsellcash,0) %>
	<%
		'할인가
		if oitem.FItemList(i).Fsailyn="Y" then
			Response.Write "<br><font color=#F08050>("&CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%할)" & FormatNumber(oitem.FItemList(i).Fsailprice,0) & "</font>"
		end if
		'쿠폰가
		if oitem.FItemList(i).FitemCouponYn="Y" then
			Select Case oitem.FItemList(i).FitemCouponType
				Case "1"
					Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
				Case "2"
					Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(oitem.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
			end Select
		end if
	%>
	</td>
	<td align="left"><%= oItem.FItemList(i).FMakerID %></td>
	<td>
	<% if oItem.FItemList(i).IsUpcheBeasong() then Response.Write "업체배송":Else Response.Write "10X10" %>
	</td>
	<td>
	<% if oItem.FItemList(i).FSellYn="Y" then %>
	Y
	<% else %>
	N
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center">
	<% if oItem.HasPreScroll then %>
		<a href="javascript:goPage(<%= oItem.StartScrollPage-1 %>)">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oItem.StartScrollPage to oItem.FScrollCount + oItem.StartScrollPage - 1 %>
		<% if i>oItem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage(<%= i %>)">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oItem.HasNextScroll then %>
		<a href="javascript:goPage(<%= i %>)">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set oItem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->