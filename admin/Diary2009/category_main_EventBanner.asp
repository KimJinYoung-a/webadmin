<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� 2008����Ʈ�����̵� 2009������ ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/Diary2009/Classes/category_main_EventBannerCls.asp"-->

<%
'// ���� ����
dim cdl, page, isusing, evtCd , cdm
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	evtCd = request("evtCd")
	cdm		= request("cdm")

	if page="" then page=1
	if isusing="" then isusing="Y"

dim omd
	set omd = New CateEventBanner
	omd.FCurrPage = page
	omd.FPageSize=8
	omd.FRectCDL = cdl
	omd.FRectcdm = cdm
	omd.FRectEvtCD = evtCd
	omd.FRectIsusing = isusing
	omd.GetEventBannerList

dim i
%>

<script language='javascript'>

// ��ü üũ/����
function ckAll(){
	if(frm.idxArrTmp.length){
		for(i=0;i<frm.idxArrTmp.length;i++) {
			frm.idxArrTmp[i].checked=frm.ckall.checked;
		}
	}
	else {
		frm.idxArrTmp.checked=frm.ckall.checked;
	}
}

// ���� üũ
function CheckSelected(selc){
	if(frm.ckall.checked) {
		frm.ckall.checked=false;
		ckAll()
		selc.checked=true;
	}
}

// ���� �������� Ȯ��
function delitems(){
	var chk=0;
	if(frm.idxArrTmp.length) {
		for(i=0;i<frm.idxArrTmp.length;i++) {
			if(frm.idxArrTmp[i].checked)
				chk++;
		}
	}
	else {
		if(frm.idxArrTmp.checked)
			chk++;
	}

	if (chk==0){
		alert('���þ������� �����ϴ�.');
		return;
	}
	
	
	if (confirm('���� �������� �����Ͻðڽ��ϱ�?')){
		frm.mode.value="del";
		frm.action="doMainEventBanner.asp";
		frm.submit();
	}
}


// ��ü ������� ����
function changeUsing(upfrm){
	var chk=0;
	if(frm.idxArrTmp.length) {
		for(i=0;i<frm.idxArrTmp.length;i++) {
			if(frm.idxArrTmp[i].checked)
				chk++;
		}
	}
	else {
		if(frm.idxArrTmp.checked)
			chk++;
	}

	if (chk==0){
		alert('���þ������� �����ϴ�.');
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('���� �������� ��������� �����մϴ�');
	} else {
		var ret = confirm('���� �������� ������ ����  �����մϴ�');
	}
	
	if (ret){
		
		upfrm.mode.value="changeUsing";
		upfrm.action="doMainEventBanner.asp";
		upfrm.submit();

	}
}

// ����ä�� ���ϻ���
function delCategoryEventBanner(){
    if(!frm.cdl.value) {
    	alert("������ ī�װ��� �������ֽʽÿ�.");
    }
   	else {
	    if (confirm('�����Ͻðڽ��ϱ�?')){
			 var popwin = window.open('','refreshFrm','');
			 popwin.focus();
			 refreshFrm.target = "refreshFrm";
			 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner_del.asp?cdl=" + frm.cdl.value+"&cdm="+frm.cdm.value;
			 refreshFrm.submit();
	    }
	}
}

// �̺�Ʈ ��� ������ ���뿩�� Ȯ��
function RefreshCategoryEventBanner(){
    if(!frm.cdl.value) {
    	alert("������ ī�װ��� �������ֽʽÿ�.");
    }
	else if (frm.cdl.value == '110'){
		if (frm.cdm.value==''){
			alert('����ä���� ��ī�װ��� �����ؾ߸� �մϴ�');			
			return;
		}else{
		    if (confirm('����ä�� �����Ͻðڽ��ϱ�?')){
				 var popwin = window.open('','refreshFrm','');
				 popwin.focus();
				 refreshFrm.target = "refreshFrm";
				 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner.asp?cdl=" + frm.cdl.value + "&cdm=" + frm.cdm.value;
				 refreshFrm.submit();
		    }		
		}
	}
   	else {
	    if (confirm('�����Ͻðڽ��ϱ�?')){
			 var popwin = window.open('','refreshFrm','');
			 popwin.focus();
			 refreshFrm.target = "refreshFrm";
			 refreshFrm.action = "<%=wwwURL%>/chtml/diary/make_category_main_EventBanner.asp?cdl=" + frm.cdl.value + "&cdm=" + frm.cdm.value;
			 refreshFrm.submit();
	    }
	}
}

// ���� ����
function viewPage(idx)
{
	frm.mode.value="edit";
	frm.page.value=<%=page%>;
	frm.idx.value=idx;
	frm.action="category_main_EventBanner_input.asp";
	frm.submit();
}

function changecontent()
{
	document.frm.submit();

}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="post"></form>
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="mode" value="">
<input type="hidden" name="evtid" value="">
<input type="hidden" name="idx" value="">
<input type="hidden" name="idxarr" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			ī�װ����� : 
<select class='select' name="cdl">
<option value='010' selected>�����ι���</option>
</select>
<select class='select' name="cdm">
<option value='' <% If cdm = "" Then Response.Write "selected" End If %>>��ü���̾</option>
<option value='10' <% If cdm = "10" Then Response.Write "selected" End If %>>���ô��̾</option>
<option value='20' <% If cdm = "20" Then Response.Write "selected" End If %>>�Ϸ���Ʈ���̾</option>
<option value='30' <% If cdm = "30" Then Response.Write "selected" End If %>>ĳ���ʹ��̾</option>
<option value='40' <% If cdm = "40" Then Response.Write "selected" End If %>>������̾</option>
</select>
			<% if cdl="110" then DrawSelectBoxCategoryMid "cdm", cdl, cdm %>
			<a href="javascript:RefreshCategoryEventBanner()"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle">�Ǽ�������</a> 
			<% if cdl="110" then %>
				<input type="button" value="�Ǽ������ϻ���" onclick="delCategoryEventBanner()" class="button">
			<% end if %>
			<br>������� : <select name="isusing"><option value="Y">Yes</option><option value="N">No</option></select>
			�̺�Ʈ�ڵ� : <input type="text" name="evtCd" value="<%=evtCd%>" size="6">
			<script language="javascript">
				document.frm.isusing.value="<%=isusing%>";
			</script>			
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
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
			<select name="allusing"><option value="Y">���� -> Y</option><option value="N">���� ->N </option></select><input type="button" class="button" value="����" onclick="changeUsing(frm);">
			<% if cdl<>"" then %>
				<input type="button" value="���þ����ۻ���" onclick="delitems();" class="button">
			<% end if %>		
		</td>
		<td align="right">		
			<input type="button" value="������ �߰�" onclick="self.location='/admin/diary2009/category_main_EventBanner_input.asp?mode=add&cdl=<%= cdl %>&menupos=<%= menupos %>'" class="button">					
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if omd.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= omd.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= omd.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll()"></td>
		<td width="100" align="center">ī�װ���</td>
		<td width="150" align="center">�̺�Ʈ��</td>
		<td align="center">�̹���</td>
		<td width="50" align="center">ǥ�ü���</td>
		<td width="50" align="center">�������</td>
		<td width="80" align="center">�����</td>
    </tr>
	<% for i=0 to omd.FResultCount-1 %>
		
    <% if omd.FItemList(i).fisusing = "Y" then %>
	    <tr align="center" bgcolor="#FFFFFF">
	    <% else %>    
	    <tr align="center" bgcolor="#FFFFaa">
		<% end if %>
		<td align="center"><input type="checkbox" name="idxArrTmp" value="<%= omd.FItemList(i).fidx %>" onclick="CheckSelected(this)"></td>
		<td align="center">
			<%= omd.FItemList(i).Fcode_nm %>
			<% if omd.FItemList(i).fcdm <> "" then %>
				(<%=omd.FItemList(i).Fcode_nm_mid%>)
			<% end if %>
		</td>
		<td align="center"><a href="javascript:viewPage(<%= omd.FItemList(i).fidx %>);"><%= "[" & omd.FItemList(i).Fevt_code & "] " & omd.FItemList(i).Fevt_name %></a></td>
		<td align="center"><img src="<%= omd.FItemList(i).Fevt_bannerimg %>" width="200" border="0"></td>
		<td align="center"><%= omd.FItemList(i).Fviewidx %></td>
		<td align="center"><%= omd.FItemList(i).Fisusing %></td>
		<td align="center"><%= FormatDateTime(omd.FItemList(i).Fregdate,2) %></td>
    </tr>   

	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if omd.HasPreScroll then %>
				<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>
		
			<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
				<% if i>omd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[<%= i %>]</a>
				<% end if %>
			<% next %>
		
			<% if omd.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&evtCd=<%=evtCd%>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</form>	
</table>

<%
set omd = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->