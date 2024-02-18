<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station �̹��� & ��ũ ���� ���� ����Ʈ ������   
' History : 2009.04.01 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
dim research,isusing, fixtype, linktype, poscode, validdate
dim page

	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")

if ((research="") and (isusing="")) then 
    isusing = "Y"
    validdate = "on"
end if

if page="" then page=1

dim oposcode
set oposcode = new cposcode_list
	oposcode.FRectPosCode = poscode
	if (poscode<>"") then
	    oposcode.fposcode_oneitem
	end if

dim oMainContents
set oMainContents = new cposcode_list
	oMainContents.FPageSize = 20
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectPosCode = poscode
	oMainContents.FRectvaliddate = validdate
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

// �÷��� �Ǽ��� ����
function AssignFlashReal(upfrm,poscode,imagecount){
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
		var AssignFlashReal;
		AssignFlashReal = window.open("<%=wwwUrl%>/chtml/culturestation_mainflashmake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignFlashReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignFlashReal.focus();
}

// �̹��� �Ǽ��� ����
function AssignmaindownbarnerReal(upfrm,poscode,imagecount){
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
		var AssignmaindownbarnerReal;
		if(poscode == "500")
		{
			AssignmaindownbarnerReal = window.open("<%=wwwUrl%>/chtml/culturestation_maindownbarner_make.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignmaindownbarnerReal","width=800,height=600,scrollbars=yes,resizable=yes");
		}
		else if(poscode == "504")
		{
			AssignmaindownbarnerReal = window.open("<%=wwwUrl%>/chtml/culturestation_maindownbarner_make_2011.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignmaindownbarnerReal","width=800,height=600,scrollbars=yes,resizable=yes");
		}
		AssignmaindownbarnerReal.focus();
}

function AssignimageReal(upfrm,poscode,imagecount){
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
		var AssignimageReal;
		AssignimageReal = window.open("<%=wwwUrl%>/chtml/culturestation_imagemake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignimageReal.focus();
}

function AssignXmlReal(upfrm,poscode,imagecount){
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
		var AssignimageReal;
		AssignimageReal = window.open("<%=wwwUrl%>/chtml/culturestation_xmlmake.asp?idx=" +tot + '&poscode='+poscode+'&imagecount='+imagecount, "AssignimageReal","width=800,height=600,scrollbars=yes,resizable=yes");
		AssignimageReal.focus();
}

//���� �ڵ� ��� & ����
function popPosCodeManage(){
    var popPosCodeManage = window.open('/admin/culturestation/imagemake_poscode.asp','popPosCodeManage','width=800,height=600,scrollbars=yes,resizable=yes');
    popPosCodeManage.focus();
}

//�̹����űԵ�� & ����
function AddNewMainContents(idx){
    var AddNewMainContents = window.open('/admin/culturestation/imagemake_contents.asp?idx='+ idx,'AddNewMainContents','width=800,height=600,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}

document.domain ='10x10.co.kr';

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    <!--<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������-->
		    ��뱸��
			<select name="isusing">
			<option value="">��ü
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
			<option value="N" <% if isusing="N" then response.write "selected" %> >������
			</select>
			���뱸��
			<% call DrawMainPosCodeCombo("poscode", poscode,"") %>
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
		    <% 
		    '//���뱸�� ���ýÿ��� �Ѹ�
		    if (poscode<>"") then 
		    %>
			    <% if oposcode.FOneItem.fimagetype="flash" then %>
			    	<% if oposcode.FOneItem.fposcode = "100" then %>
			    	<a href="javascript:AssignFlashReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����</a>
			    	<% end if %>
			    <% elseif oposcode.FOneItem.fimagetype="multi" then %>
			    	<a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> 
			    	&nbsp;&nbsp;
			    	<a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
			    <% elseif oposcode.FOneItem.fimagetype="link" then %>
				    <% 
				    if oposcode.FOneItem.fposcode = "400" or oposcode.FOneItem.fposcode = "300" or oposcode.FOneItem.fposcode = "502" or oposcode.FOneItem.fposcode = "503" then 
				    %>
						<a href="javascript:AssignimageReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
			    	<% 
			    	elseif oposcode.FOneItem.fposcode = "500" or oposcode.FOneItem.fposcode = "504" then 
			    	%>
						<a href="javascript:AssignmaindownbarnerReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>			    
					<%
			    	end if
			    	%>
			    <% elseif oposcode.FOneItem.fimagetype="xml" then %>
			    	<%
			    	if oposcode.FOneItem.fposcode = "501" or oposcode.FOneItem.fposcode = "505" then
			    	%>
						<a href="javascript:AssignXmlReal(frm,<%= poscode %>,<%=oposcode.FOneItem.fimagecount%>);"><img src="/images/refreshcpage.gif" border="0"> XML Real ����</a>			    
			    	<%
			    	end if
			    end if 
			    %>

		    <% end if %>
		</td>
		<td align="right">
			<% if C_ADMIN_AUTH then %>
			<input type="button" value="�ڵ����" class="button" onClick="popPosCodeManage();">
			<% end if %>
		
			<input type="button" value="�űԵ��" class="button" onClick="javascript:AddNewMainContents('0');">						
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
	    <td align="center">Idx</td>
	    <td align="center">Image</td>
	    <td align="center">���и�</td>
	    <td align="center">LinkType</td>
	    <td align="center">�켱����</td>
	    <td align="center">��뿩��</td>
	    <td align="center">�����</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			 		
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr bgcolor="#DDDDDD">
		<% else %>
			<tr bgcolor="#FFFFFF">
		<% end if %>	
		<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>		
	    <td align="center"><%= oMainContents.FItemList(i).Fidx %><input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>"></td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<img width=40 height=40 src="<%=uploadUrl%>/culturestation/main/<%= oMainContents.FItemList(i).fimagepath %>" border="0">
	    	</a>
	    </td>
	    <td align="center">
	    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');">
	    	<%= oMainContents.FItemList(i).Fposname %>
	    	(<%
	    	if oMainContents.FItemList(i).fitemid <> 0 and oMainContents.FItemList(i).fitemid <> "" then 
	    	response.write oMainContents.FItemList(i).fitemid
	    	elseif oMainContents.FItemList(i).fevt_code <> 0 and oMainContents.FItemList(i).fevt_code <> "" then 
	    	response.write oMainContents.FItemList(i).fevt_code
	    	end if
	    	%>)</a>
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fimagetype %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fimage_order %></td>
	    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fregdate %></td> 
	</tr>
	</form>	
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="7" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oMainContents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

		


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

