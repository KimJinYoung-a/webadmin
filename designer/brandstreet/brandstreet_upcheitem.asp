<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü �귣�������� ���� 
' History : 2009.03.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/brandstreet/brandstreet_upche_cls.asp"-->

<%
dim page , itemid , types
	types   = requestCheckvar(request("types"),32)
	itemid  = requestCheckvar(request("itemid"),10)
	page    = requestCheckvar(request("page"),10)
	
if page="" then page=1
if types = "" then types = 1
dim oMainContents
set oMainContents = new cbrandstreet_list
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.frectmakerid = session("ssBctId")	
	oMainContents.frectitemid = itemid
	oMainContents.fupche_item

dim i

%>

<script language="javascript">

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}	

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

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

//����
function item_process(upfrm){
	var type
	if (upfrm.types.selectedIndex==0){
		alert('���뱸���� �����ϼ���');
		upfrm.types.focus();
		return;
	}
	type=upfrm.types.value;
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
					
						upfrm.fidx.value = upfrm.fidx.value + frm.itemid.value + "," ;								
					}
				}
			}

	submit_frm.itemid.value = upfrm.fidx.value;
	upfrm.fidx.value = ""
	submit_frm.type.value = upfrm.types.value;	
							
	submit_frm.action= '/designer/brandstreet/brandstreet_upcheitem_process.asp';
	submit_frm.submit();
}

</script>
<form name="submit_frm" method="post">
	<input type="hidden" name="itemid">
	<input type="hidden" name="type">
</form>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			��ǰ�ڵ� : <input type="text" value="<%=itemid %>" name="itemid" size="10">

		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>

</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		���Ǹ����� ��ǰ�� �˻� �˴ϴ�.
		</td>
		<td align="right">
			���뱸��
			<select name="types">
			<option value="">��ü</option>
			<option value="1" <% if types="1" then response.write "selected" %>>�ߴܹ��</option>
			</select>
		
			<input type="button" onclick="item_process(frm);" value="����" class="button" >		
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oMainContents.FResultCount > 0 then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oMainContents.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
 		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	    <td align="center">Image</td>
	    <td align="center">��ǰ�ڵ�</td>
	    <td align="center">��ǰ��</td>
	    <td align="center">�Ǹſ���</td>
	      
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for�� �ȿ��� i ���� ������ ����-->	 		
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr bgcolor="#DDDDDD">
		<% else %>
			<tr bgcolor="#FFFFFF">
		<% end if %>	
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
			<input type="hidden" name="itemid" value="<%= oMainContents.FItemList(i).fitemid %>">
		</td>		
	    <td align="center">
	    	<img width=40 height=40 src="<%= oMainContents.FItemList(i).fsmallimage %>" border="0">
	    </td>
	    <td align="center">
	    	<%= oMainContents.FItemList(i).fitemid %>
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fitemname %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fsellyn %></td>
	    
	</tr>
	</form>	
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oMainContents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>&itemid=<%=itemid%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&itemid=<%=itemid%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&itemid=<%=itemid%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

