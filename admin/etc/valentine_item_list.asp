<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/valentinecls.asp"-->
<%
dim page, idx, ovalen,masterid
page = request("page")
if page="" then page=1

idx = request("idx")
masterid = request("masterid")
if masterid="" then masterid=1

set ovalen = new ValentineItem
ovalen.FPageSize = 20
ovalen.FCurrPage = page
ovalen.GetValentineItemList masterid

dim i

Sub SelectMaster(byval selectedId)
   dim tmp_str,query1
   %><select name="masterid" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select idx,title from [db_contents].[dbo].tbl_blood_master"
   query1 = query1 & " order by idx Asc"

   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("idx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&rsget("title")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub
%>


<script language='javascript'>
function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function addnewItem(frm){
	var popwin;
	popwin = window.open("/admin/pop/viewitemlist.asp?designerid=" + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
	popwin.focus();
}

function delthis(frm){
	var ret= confirm('삭제 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function arrsave(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('선택 아이템이 없습니다.');
		return;
	}

	upfrm.idxarr.value = "";
	upfrm.viewidxarr.value = "";
	upfrm.itemidarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (!IsDigit(frm.viewidx.value)){
				alert('표시순서는 숫자만 가능합니다.');
				frm.viewidx.focus();
				return;
			}

			upfrm.idxarr.value = upfrm.idxarr.value + frm.idx.value + "|";
			upfrm.viewidxarr.value = upfrm.viewidxarr.value + frm.viewidx.value + "|";
			upfrm.itemidarr.value = upfrm.itemidarr.value + frm.itemid.value + "|";
		}
	}

	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		upfrm.submit();
	}
}

function AddIttems(){
	frmbuf.submit();
}

function QuickAdd(frm){

	frmbuf.itemidarr.value = frm.txitemarr.value;

	if (frmbuf.itemidarr.value.length<1){
		alert('값을 입력하세요.');
		return;
	}

	var ret=confirm('저장하시겠습니까?');

	if(ret){
		frmbuf.submit();
	}
}

function changecontent(){
	document.frm.submit();
}
</script>


<table width="800" cellspacing="1" class="a" >
<form name="frm">
<tr>
	<td align="left"><% SelectMaster masterid %></td>
</tr>
</form>
</table>
<table width="800" cellspacing="1" class="a" >
<form name="frmquick" >
<tr>
	<td colspan="2">
		<input type="text" name="txitemarr" size="90" value=""><input type="button" value="ItemID로 등록" onClick="QuickAdd(frmquick);">
	</td>
</tr>
</form>
<tr>
	<td align="left"></td>
	<td align="right"><input type="button" value="새상품 추가" onclick="addnewItem('frmbuf.itemidarr')"></td>
</tr>
</table>
<table width="800" cellspacing="1" class="a" bgcolor="#3d3d3d">
    <tr bgcolor="#DDDDFF">
    	<td width="70">상품ID</td>
    	<td width="50">이미지</td>
    	<td width="120">상품명</td>
		<td width="50">삭제</td>
    </tr>
    <% for i=0 to ovalen.FResultCount -1 %>
    <form name="frmBuyPrc_<%= i %>" method="post" action="dovalentineedit.asp">
    <input type="hidden" name="mode" value="deleventdetail">
    <input type="hidden" name="idx" value="<%= ovalen.FItemList(i).Fidx %>">
    <input type="hidden" name="itemid" value="<%= ovalen.FItemList(i).Fitemid %>">
    <tr bgcolor="#FFFFFF">
    	<td><%= ovalen.FItemList(i).FItemId %></td>
    	<td><img src="<%= ovalen.FItemList(i).FImageSmall %>" width="50" height="50"></td>
    	<td><%= ovalen.FItemList(i).FItemName %></td>
    	<td><a href="javascript:delthis(frmBuyPrc_<%= i %>)">삭제</a></td>
    </tr>
    </form>
    <% next %>
    <tr bgcolor="#FFFFFF">
		<td colspan="7" align="center">
		<% if ovalen.HasPreScroll then %>
			<a href="?page=<%= ovalen.StarScrollPage-1 %>&menupos=<%= menupos %>&idx=<%=idx%>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ovalen.StarScrollPage to ovalen.FScrollCount + ovalen.StarScrollPage - 1 %>
			<% if i>ovalen.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&menupos=<%= menupos %>&idx=<%=idx%>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ovalen.HasNextScroll then %>
			<a href="?page=<%= i %>&menupos=<%= menupos %>&idx=<%=idx%>">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<form name="frmbuf" method="post" action="dovalentineedit.asp">
<input type="hidden" name="mode" value="addeventdetailarr">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="masterid" value="<% = masterid %>">
</form>

<form name="frmArrupdate" method="post" action="dovalentineedit.asp">
<input type="hidden" name="mode" value="modieventdetail">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="viewidxarr" value="">
<input type="hidden" name="masterid" value="<% = masterid %>">
<input type="hidden" name="itemidarr" value="">
</form>
<%
set ovalen = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->