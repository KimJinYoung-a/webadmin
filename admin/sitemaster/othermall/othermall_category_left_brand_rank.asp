<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2007.11.09 �ѿ�� ����
'			2008.06.18 �ѿ�� ����/�߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/othermall_idx_mdchoice_brandcls.asp"-->
<%

Sub SelectBoxDesignerItem()
   dim query1,tmp_str
%>
	<select name="makerid">
		<option value=''>-- ��ü���� --</option>
<%
			query1 = " select userid,socname_kor,defaultmargine from [db_user].dbo.tbl_user_c"
			rsget.Open query1,dbget,1

			if  not rsget.EOF  then
				rsget.Movefirst
				
				do until rsget.EOF
				   response.write("<option value='"&rsget("userid")& "'>" & rsget("userid") & "  [" & replace(db2html(rsget("socname_kor")),"'","") & "]" & "</option>")
				   rsget.MoveNext
				loop
			end if
			
			rsget.close
			response.write("</select>")
End Sub

dim cdl, page
cdl = request("cdl")
page = request("page")

if page="" then page=1

dim omd
set omd = New MDChoice
omd.FCurrPage = page
omd.FPageSize=100
omd.FRectCDL = cdl
omd.GetCategoryLeftBrandRank

dim i
%>

<script language='javascript'>

function popItemWindow(iid,frm){
	window.open("/admin/pop/viewitemlist.asp?designerid=" + iid + "&target=" + frm, "popup_item", "width=800,height=500,scrollbars=yes,status=no");
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

function delitems(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}

		upfrm.mode.value="del";
		upfrm.action="othermall_doleftbrandrank.asp";
		upfrm.submit();

	}
}

function AddIttems2(){
	
	if (frm.makerid.value==""){
		alert("�귣�带 �������ּ���!");
		return;
	}
	if (frm.cdl.value==""){
		alert("ī�װ��� �������ּ���!");
		return;
	}
	var ret = confirm('�߰��Ͻðڽ��ϱ�?');
	
	if (ret){
		frm.mode.value="add";
		frm.action = "othermall_doleftbrandrank.asp";
		frm.submit();
	}
}

function RefreshBestBrand(upfrm){

	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	var ret = confirm('10������ ����˴ϴ�. ���� �������� ���� �Ͻðڽ��ϱ�?');

	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "," ;
				}
			}
		}

		//upfrm.mode.value="del";
		upfrm.action = "<%=othermall%>/chtml/othermall_make_best_friend.asp"
		upfrm.submit();

	}
}

function changecontent(){
    // nothing
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">���<br>����</td>
		<td align="left">   		
		</td>		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<!--<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">-->
			<input type="button" value="����ϱ�" onclick="AddIttems2()" class="button">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<% DrawSelectBoxCategoryLarge "cdl", cdl %>&nbsp;<% SelectBoxDesignerItem %>
			<input type="hidden" name="mode"> 	
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			�������� ��ų  �����ۼ����� <a href="javascript:RefreshBestBrand(refreshFrm);">
		    <img src="/images/refreshcpage.gif" width="19" align="absmiddle" border="0"></a> ��ư�� �����ּ���
		</td>
		<td align="right">
			<input type="button" value="���þ����� ����" onClick="delitems(delform)" class="button">	
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if omd.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= omd.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= omd.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="150" align="center">idx</td>
		<td width="150" align="center">ī�װ���</td>
		<td width="200" align="center">��ü��</td>
		<td width="150" align="center">�̹���</td>
    </tr>
		<% for i=0 to omd.FresultCount-1 %>
		<form name="frmBuyPrc_<%=i%>" method="post" action="" >
		<input type="hidden" name="itemid" value="<%= omd.FItemList(i).Fidx %>">
	    <tr align="center" bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td align="center"><%= omd.FItemList(i).Fidx %></td>
			<td align="center"><%= omd.FItemList(i).GetCD1Name %></td>
			<td align="center"><%= omd.FItemList(i).Fmakerid %></td>
			<td align="center"><img src="<%= omd.FItemList(i).FImgSmall %>"><img src="<%= omd.FItemList(i).Ftitleimgurl %>" ></td>
	    </tr>   
		</form>
		<% next %>
		
	<% else %>
	
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
		
	<% end if %>
	
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omd.HasPreScroll then %>
				<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
				<% if i>omd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if omd.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&cdl=<%=cdl%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<form name="delform" method="post" action="doleftbrandrank.asp">
<input type="hidden" name="mode">
<input type="hidden" name="itemid">
</form>
<form name="refreshFrm" method=post>
<input type="hidden" name="cdl">
<input type="hidden" name="itemid">
</form>
<%
set omd = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->