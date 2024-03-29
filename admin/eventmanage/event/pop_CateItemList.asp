<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
'###############################################
' PageName : pop_CateItemList.asp
' Discription : 카테고리별 상품 목록 팝업
'	- 받은 카테고리의 상품을 보여주고 선택한 상품코드를 반환
' History : 2008.04.01 허진원 : 생성
'			2008.06.11 정윤정 수정 : 정렬방식 선택
'###############################################

dim makerid,itemid,itemname
dim page,sellyn,packyn
dim target
dim cdl,cdm,cds
dim sortdiv

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

makerid = request("makerid")
itemid = request("itemid")
itemname = request("itemname")
sellyn = request("sellyn")
page = request("page")
target= request("target")
sortdiv =request("sd")
if page="" then page=1

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
oItem.FRectSortDiv = sortdiv
oItem.GetItemList

dim i
%>
<script>
function SelectItems(){
	var tg = eval('opener.frmWin.itemuse');
	var tgvalue="";
	var tgid = eval('opener.frmWin.itemuse_itemid');
	var tgidvalue="";

	var frm;
	var pass = false;
	var temp = 0;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('선택 상품이 없습니다.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				tgvalue = tgvalue + "[" + frm.itemid.value + "]" + frm.itemname.value + " " ;
				tgidvalue = tgidvalue + frm.itemid.value + " " ;
				temp = temp + 1;
			}
		}
	}

	if (temp > 1)
	{
		alert("한개만 선택이 가능합니다.");
		return;
	}

	tg.focus();
	tg.value = tgvalue;
	tgid.value = tgidvalue;
	opener.ViewByteLength();
	window.close();

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
				<font color="#333333"><b>카테고리 상품 선택/추가</b></font>
			</td>
			<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
		카테고리에 속한 상품을 검색/선택합니다.
	</td>
</tr>
<tr><td height="10"></td></tr>
</table>
<!-- 검색 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="sd" value="<%=sortdiv%>">
<input type="hidden" name="target" value="<%= target %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
			디자이너 :
			<% drawSelectBoxDesigner "makerid",makerid %>
			상품ID :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="9">
			상품명 :
			<input type="text" class="text" name="itemname" value="<%= itemname %>" size="8" maxlength="32">
			판매여부 :
			<select name="sellyn" class="select">
		                 	<option value='' selected>선택</option>
		                 	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
		                 	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
		         	</select>
			</td>
		</tr>
		<tr>
			<td>
				<!-- #include virtual="/common/module/categoryselectbox.asp"-->
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
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<form name="frmttl" onsubmit="return false;">
<tr  style="padding:10 0 10 0;">
	<td>
		<!--<input type="button" class="button" value="전체선택" onClick="AnSelectAllFrame(true)">//-->
	</td>
	<td align="right">
		<input type="button" class="button" value="상품선택" onClick="SelectItems()">
	</td>
</tr>
</form>
</table>
<!-- 본문 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		검색결과 : <b><%= FormatNumber(oItem.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b><%= FormatNumber(page,0) %> / <%= FormatNumber(oItem.FTotalPage,0) %></b>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<font color="red"><b>※ 반드시 상품을 한개만 선택하세요.</b></font>
	</td>
	<td colspan="2" align="right">정렬:
		<select name="sortDiv" onchange="frm.sd.value = this.value; frm.submit();">
		<option value="new" <% IF sortDiv="new" Then response.write "selected" %> >신상품순</option>
		<option value="cashH" <% IF sortDiv="cashH" Then response.write "selected" %>>높은가격순</option>
		<option value="cashL" <% IF sortDiv="cashL" Then response.write "selected" %>>낮은가격순</option>
		<option value="best" <% IF sortDiv="best" Then response.write "selected" %>>베스트순</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">선택</td>
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
<input type="hidden" name="itemname" value="<%= oItem.FItemList(i).FItemName %>">
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="radio" name="cksel" onClick="AnCheckClick(this);"></td>
	<td><%= oItem.FItemList(i).FItemID %></td>
	<td><img src="<%= oItem.FItemList(i).Fsmallimage %>" width="50" height="50" border=0 alt=""></td>
	<td align="left"><%= oItem.FItemList(i).FItemName %></td>
	<td align="left"><%= FormatNumber(oItem.FItemList(i).Fsellcash,0) %></td>
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
	<td colspan="9" align="center">
	<% if oItem.HasPreScroll then %>
		<a href="?page=<%= oItem.StartScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oItem.StartScrollPage to oItem.FScrollCount + oItem.StartScrollPage - 1 %>
		<% if i>oItem.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oItem.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&makerid=<%= makerid %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cdl=<% = cdl %>&cdm=<% = cdm %>&cds=<% = cds %>&sd=<%=sortdiv%>">[next]</a>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->