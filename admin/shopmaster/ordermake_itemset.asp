<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/buypricecls.asp"-->
<%
response.write "사용중지"
dbget.close()	:	response.End

dim makerid,itemid,itemname
dim page,dispyn,sellyn,usingyn,packyn
dim mduserid
makerid = request("makerid")
itemid = request("itemid")
itemname = request("itemname")
dispyn = request("dispyn")
sellyn = request("sellyn")
usingyn = request("usingyn")
packyn = request("packyn")
page = request("page")
mduserid = request("mduserid")

if page="" then page=1

dim obuyprice
set obuyprice = new CBuyPrice
obuyprice.FCurrPage = page
obuyprice.FPageSize = 50
obuyprice.FSearchItemName = itemname
obuyprice.FSearchDesigner = makerid
obuyprice.FSearchItemid = itemid
obuyprice.FSearchDispYn = dispyn
obuyprice.FSearchSellYn = sellyn
obuyprice.FSearchusingyn = usingyn
obuyprice.getPrcList

dim i
%>

<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function AllSelectOBJ(bool,obj){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (eval("frm." + obj + "[1].disabled!=true")){
				eval("frm." + obj + "[1].checked = true");
			}
		}
	}
}

function CheckNSelectOBJ(bool,obj){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert("선택 아이템이 없습니다.");
		return;
	}
	//alert(bool);
	//return;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				 eval("frm." + obj + "[1].checked = true");
			}
		}
	}
}

function CheckYSelectOBJ(bool,obj){
	var frm;
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert("선택 아이템이 없습니다.");
		return;
	}
	//alert(bool);
	//return;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				 eval("frm." + obj + "[0].checked = true");
			}
		}
	}
}

function ChangeOrderMakerFrame(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

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

	var ret = confirm('선택 상품을 저장하시겠습니까?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;

					if (frm.itemdiv[0].checked){
						upfrm.itemdiv.value = upfrm.itemdiv.value + "|" + "01";
					}
					else if (frm.itemdiv[1].checked){
						upfrm.itemdiv.value = upfrm.itemdiv.value + "|" + "06";
					}
					else{
						upfrm.itemdiv.value = upfrm.itemdiv.value + "|" + "88";
					}

				}
			}
		}
		frm.submit();
	}
}
</script>


<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
		브랜드:
		<% drawSelectBoxDesigner "makerid",makerid %>
		<br>
		상품ID:
		<input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
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
             	사용여부:
		<select name="usingyn">
                     	<option value='' selected>선택</option>
                     	<option value='Y' <% if usingyn="Y" then response.write "selected" %> >Y</option>
                     	<option value='N' <% if usingyn="N" then response.write "selected" %> >N</option>
             	</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#000000">
<tr bgcolor="FFFFFF">
	<td colspan="10" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(obuyprice.FTotalPage,0) %> count: <%= FormatNumber(obuyprice.FTotalCount,0) %></td>
</tr>
<tr bgcolor="FFFFFF">
	<form name="frmttl" onsubmit="return false;">
	<td colspan="10" height="30">
	 <table width="100%">
		  <tr>
			 <td align="left"><input type="button" value="전체선택" onClick="AnSelectAllFrame(true)">&nbsp;<input type="button" value="선택상품저장" onClick="ChangeOrderMakerFrame()"></td>
			 <td align="right"><input type="button" value="모두주문제작상품으로 선택" onClick="AllSelectOBJ('06','itemdiv')"></td>
		  </tr>
	 </table>
	</td>
	</form>
</tr>
<tr bgcolor="DDDDFF" align="center">
	<td>선택</td>
	<td width="50">이미지</td>
	<td>상품ID</td>
	<td>상품명</td>
	<td>브랜드</td>
	<td>배송구분</td>
	<td width="80">상품구분</td>
</tr>
<% for i=0 to obuyprice.FresultCount-1 %>
<form name="frmBuyPrc_<%= obuyprice.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="doitemviewset.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= obuyprice.FItemList(i).FItemID %>">
<tr align="center" bgcolor="FFFFFF">
	<% if obuyprice.FItemList(i).Fitemdiv="88" or obuyprice.FItemList(i).Fitemdiv="90" then %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" disabled></td>
	<% else %>
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"  ></td>
	<% end if %>
	<td><img src="<%= obuyprice.FItemList(i).FImageSmall %>" width="50" height="50"></td>
	<td><a href="javascript:PopItemSellEdit('<%= obuyprice.FItemList(i).FItemID %>');"><%= obuyprice.FItemList(i).FItemID %></a></td>
	<td><%= obuyprice.FItemList(i).FItemName %></td>
	<td><%= obuyprice.FItemList(i).FMakerID %></td>
	<td>
	<% if obuyprice.FItemList(i).FBaesongGB="1" then %>
		10x10
	<% else %>
	   	<font color=red><%= BaesongCd2Name(obuyprice.FItemList(i).FBaesongGB) %></font>
	<% end if %>
	</td>
	<td>
	<input type="radio" name="itemdiv" value="01" <% if obuyprice.FItemList(i).Fitemdiv="01" then response.write "checked" %>>일반상품
	<input type="radio" name="itemdiv" value="06" <% if obuyprice.FItemList(i).Fitemdiv="06" then response.write "checked" %>>주문제작상품
	<!-- <input type="radio" name="itemdiv" value="88" <% if obuyprice.FItemList(i).Fitemdiv="88" then response.write "checked" %> >DIY상품 -->
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="FFFFFF">
	<td colspan="10" align="center">
	<% if obuyprice.HasPreScroll then %>
		<a href="javascript:NextPage('<%= obuyprice.StarScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + obuyprice.StarScrollPage to obuyprice.FScrollCount + obuyprice.StarScrollPage - 1 %>
		<% if i>obuyprice.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if obuyprice.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>

<form name="frmArrupdate" method="post" action="doitemmakerset.asp">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemdiv" value="">
</form>
</table>
<%
set obuyprice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->