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
dim page , isusing , types
	types = requestCheckVar(request("types"),50)
	isusing = requestCheckVar(request("isusing"),30)
	page    = requestCheckVar(request("page"),10)
	
if page="" then page=1
if types = "" then types = 1
if isusing = "" then isusing = "Y"
dim oMainContents
set oMainContents = new cbrandstreet_list
	oMainContents.FPageSize = 6
	oMainContents.FCurrPage = page
	oMainContents.frectisusing = isusing
	oMainContents.frecttype = types	
	oMainContents.frectmakerid = session("ssBctId")		
	oMainContents.fcontents_list

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

// �ϰ� ������
function display_no(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
					
				}
			}
		}
	var tot;
	tot = upfrm.fidx.value;
	upfrm.fidx.value = ""
	var display_no;
	display_no = window.open("/designer/brandstreet/brandstreet_process.asp?itemid=" +tot + '&mode=isusing_no', "display_no","width=400,height=300,scrollbars=yes,resizable=yes");
	display_no.focus();
}


//�űԵ�� 
function AddNewMainContents(){
    var AddNewMainContents = window.open('/designer/brandstreet/brandstreet_upcheitem.asp','AddNewMainContents','width=600,height=768,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}


</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			���뱸��
			<select name="types">
			<option value="">��ü</option>
			<option value="1" <% if types="1" then response.write "selected" %>>�ߴܹ��</option>
			</select>
		    ���⿩��
			<select name="isusing">
			<option value="">��ü</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����</option>
			<option value="N" <% if isusing="N" then response.write "selected" %> >������</option>
			</select>

		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" onclick="display_no(frm);" value="�ϰ��������" class="button" >	
		</td>
		<td align="right">		
			<input type="button" onclick="AddNewMainContents();" value="�űԵ��" class="button" >		
		</td>
	</tr>
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
	    <td align="center">���и�</td>
	    <td align="center">��ǰ�ڵ�</td>
	    <td align="center">��ǰ��</td>
	    <td align="center">���⿩��</td>
	    
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
			<input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>">
		</td>		
	    <td align="center">
	    	<img width=40 height=40 src="<%= oMainContents.FItemList(i).fsmallimage %>" border="0">
	    </td>
	    <td align="center">
	    	<%
	    	if oMainContents.FItemList(i).ftype = "1" then
	    		response.write "�ߴܹ��"
	    	end if 
	    	%>
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fitemid %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fitemname %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fisusing %></td>
	    
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
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>&isusing=<%=isusing%>&types=<%=types%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&types=<%=types%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&types=<%=types%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

