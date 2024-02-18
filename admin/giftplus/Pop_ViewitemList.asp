<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/buypricecls.asp"-->
<%
dim designerid,itemid,itemname
dim page,dispyn,sellyn,packyn
dim target,gubun
dim cd1,cd2,cd3

cd1 = request("cd1")
cd2 = request("cd2")
cd3 = request("cd3")

designerid = request("designerid")
itemid = request("itemid")
itemname = request("itemname")
dispyn = request("dispyn")
sellyn = request("sellyn")
page = request("page")
target= request("target")
gubun= request("gubun")

if page="" then page=1

dim obuyprice
set obuyprice = new CBuyPrice
obuyprice.FCurrPage = page
obuyprice.FPageSize = 30
obuyprice.FSearchItemName = itemname
obuyprice.FSearchDesigner = designerid
obuyprice.FSearchItemid = itemid
obuyprice.FSearchDispYn = dispyn
obuyprice.FSearchSellYn = sellyn
obuyprice.FRectCD1 = cd1
obuyprice.FRectCD2 = cd2
obuyprice.FRectCD3 = cd3
obuyprice.getPrcList

dim i
%>
<script>
function SelectItems(){
	var tg = eval('opener.addfrm.arrItemID');
	var tgvalue="";

	var chkbx = document.getElementsByName('cksel');

	for (var i=0;i<chkbx.length;i++) {
			if (chkbx[i].checked){
				tgvalue=tgvalue  + chkbx[i].value + ",";
			}
	}

	if (tgvalue.length < 1){
			alert('하나 이상 상품을 선택해 주세요');
			return;
	} else {

		tg.value = tgvalue;
		opener.AddItems();
	}
	self.close();
}

function changecontent(){
	document.frm.submit();
}
function ManualSelectItems(){
	var tg = eval('opener.addfrm.arrItemID');
	var tgvalue=document.getElementById('arritemid');
	tgvalue=tgvalue.value;



	if (tgvalue.length < 1){
		alert('상품코드를 입력해 주세요');
		return;
	} else {

		if (tgvalue.charAt((tgvalue.length)-1)!=','){
			tgvalue = tgvalue + ',' ;
		}

		tg.value = tgvalue;
		opener.AddItems();
	}
}


function AnSelectAllChk(bool){
	var frm = document.getElementsByName('cksel');
	for (var i=0;i<frm.length;i++){
		if (frm[i].disabled!=true){
			frm[i].checked = bool;
		AnCheckClick(frm[i]);
		}
	}
}

</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="target" value="<%= target %>">
	<tr>
		<td class="a" >
		디자이너:
		<% drawSelectBoxDesigner "designerid",designerid %>
		상품ID:
		<input type="text" name="itemid" value="<%= itemid %>" size="10" maxlength="9">
		상품명:
		<input type="text" name="itemname" value="<%= itemname %>" size="8" maxlength="32">
		전시여부:
		<select name="dispyn">
                     	<option value='' selected>선택</option>
                     	<option value='Y' <% if dispyn="Y" then response.write "selected" %> >Y</option>
                     	<option value='N' <% if dispyn="N" then response.write "selected" %> >N</option>
             	</select>
		판매여부:
		<select name="sellyn">
                     	<option value='' selected>선택</option>
                     	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
                     	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
             	</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="a">
			카테고리 선택 : <% DrawSelectBoxCategoryLarge "cd1", cd1 %>
			<% DrawSelectBoxCategoryMid "cd2", cd1, cd2 %>
			<% DrawSelectBoxCategorySmall "cd3", cd1, cd2, cd3 %>
			&nbsp;가격대별
			 <select name="gubun">
				 <option value="">선택</option>
				 <option value="01" <% if gubun = "01" then response.write "selected" %>>Price or Man or 할인</option>
				 <option value="02" <% if gubun = "02" then response.write "selected" %>>Design or Woman or 사은품</option>
				 <option value="03" <% if gubun = "03" then response.write "selected" %>>Quality or Couple or 쿠폰</option>
			 </select>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="14" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(obuyprice.FTotalPage,0) %> count: <%= FormatNumber(obuyprice.FTotalCount,0) %></td>
</tr>
<tr>
	<td colspan="14"><input type="text" name="arritemid" size="50"><input type="button" value="직접 입력" onclick="ManualSelectItems();"></td>
</tr>
<tr>
	<form name="frmttl" onsubmit="return false;">
	<td colspan="9" height="30"><input type="button" value="전체선택" onClick="AnSelectAllChk(true)">&nbsp;<input type="button" value="상품선택" onClick="SelectItems();"></td>
	</form>
</tr>
<tr>
	<td align="center">선택</td>
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">상품명</td>
	<td align="center">가격</td>
	<td align="center">디자이너</td>
	<td align="center">배송구분</td>
	<td align="center">전시여부</td>
	<td align="center">판매여부</td>
</tr>
<tr>
	<td colspan="9" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
<% for i=0 to obuyprice.FresultCount-1 %>
<tr height="20">
	<td><input type="checkbox" name="cksel" value="<%= obuyprice.FItemList(i).FItemID %>" onClick="AnCheckClick(this);"></td>
	<td><%= obuyprice.FItemList(i).FItemID %></td>
	<td><img src="<%= obuyprice.FItemList(i).FImageSmall %>" width="50" height="50" border=0 alt=""></td>
	<td><%= obuyprice.FItemList(i).FItemName %></td>
	<td><%= FormatNumber(obuyprice.FItemList(i).FSellPrice,0) %></td>
	<td><%= obuyprice.FItemList(i).FMakerID %></td>
	<td align="center">
	<% if obuyprice.FItemList(i).FBaesongGB="1" then %>
		10x10
	<% else %>
	   	<font color=red><%= BaesongCd2Name(obuyprice.FItemList(i).FBaesongGB) %></font>
	<% end if %>
	</td>
	<td align="center"><%= obuyprice.FItemList(i).FDisplayYn %></td>
	<td align="center"><%= obuyprice.FItemList(i).FSellYn %></td>
</tr>
<tr>
	<td colspan="9" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
<% next %>
<tr>
	<td colspan="9" align="center">
	<% if obuyprice.HasPreScroll then %>
		<a href="?page=<%= obuyprice.StarScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cd1=<% = cd1 %>&cd2=<% = cd2 %>&cd3=<% = cd3 %>&gubun=<% = gubun %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + obuyprice.StarScrollPage to obuyprice.FScrollCount + obuyprice.StarScrollPage - 1 %>
		<% if i>obuyprice.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cd1=<% = cd1 %>&cd2=<% = cd2 %>&cd3=<% = cd3 %>&gubun=<% = gubun %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if obuyprice.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&packyn=<%=packyn%>&target=<%= target %>&cd1=<% = cd1 %>&cd2=<% = cd2 %>&cd3=<% = cd3 %>&gubun=<% = gubun %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>

<tr>
	<td colspan="9" height="20">
</tr>
</table>
<%
set obuyprice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->