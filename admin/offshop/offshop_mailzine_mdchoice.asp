<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 오프라인 메일진
' History : 최초생성자모름
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopmailzine_mdchoicecls.asp"-->
<%
dim masteridx, page
masteridx = requestCheckVar(request("masteridx"),10)
page = requestCheckVar(request("page"),10)

if page="" then page=1

dim omd
set omd = New COffshopMailzine
omd.FCurrPage = page
omd.FPageSize=20
omd.FRectMasteridx = masteridx
omd.GetNewitemList

dim i
%>
<script language='javascript'>

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
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}

	var ret = confirm('선택 아이템을 삭제하시겠습니까?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

</script>
<form name="frmarr" method="post" action="/admin/offshop/lib/domailzinemdchoice.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
</form>
<table width="650" border="0" cellpadding="5" cellspacing="0">
<tr>
	<td><a href="/admin/offshop/offshop_mailzine_mdchoice_reg.asp"><font color="red">아이템추가</font></a></td>
</tr>
</table>
<table width="650" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		메일진구분 : <% DrawSelectBoxMailzine masteridx %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table>
<tr>
	<td><input type="button" value="선택아이템 삭제" onClick="delitems(frmarr)" class="button"></td>
</tr>
</table>
<table width="650" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF" height="25">
	<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="100" align="center">마스터구분</td>
	<td width="80" align="center">ItemID</td>
	<td width="80" align="center">Image</td>
	<td align="center">제품명</td>
	<td width="80" align="center">사용유무</td>
</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= omd.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= FormatDate(omd.FItemList(i).Fregdate,"0000.00.00") %></td>
	<td align="center"><img src="<%= omd.FItemList(i).Fimagesmall %>" width="50" height="50"></td>
	<td align="center"><%= omd.FItemList(i).FItemID %></td>
	<td align="center"><%= omd.FItemList(i).FItemname %></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>

<%
'메일진 선택
Sub DrawSelectBoxMailzine(byval selectedId)
   dim tmp_str,query1
   %><select name="masteridx" onChange="changecontent()">
     <option value="" <% if selectedId="" then response.write " selected"%>>선택</option><%
   query1 = " select idx,regdate from [db_shop].[dbo].tbl_shopmaster_mail"
   query1 = query1 + " where isusing = 'Y'"
   query1 = query1 + " order by regdate desc"
   rsget.Open query1,dbget,1

   if  not rsget.EOF  then
       rsget.Movefirst

       do until rsget.EOF
           if Cstr(selectedId) = Cstr(rsget("idx")) then
               tmp_str = " selected"
           end if
           response.write("<option value='"&rsget("idx")&"' "&tmp_str&">"&FormatDate(rsget("regdate"),"0000.00.00")&"</option>")
           tmp_str = ""
           rsget.MoveNext
       loop
   end if
   rsget.close
   response.write("</select>")

end Sub
%>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->