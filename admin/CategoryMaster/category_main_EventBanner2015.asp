<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_main_EventBannerCls.asp" -->

<%
'// ���� ����
dim cdl, page, isusing, evtCd , cdm, vCateCode
	cdl = request("cdl")
	page = request("page")
	isusing = request("isusing")
	evtCd = request("evtCd")
	cdm		= request("cdm")
	vCateCode = Request("catecode")

	if page="" then page=1
	if isusing="" then isusing="Y"

dim omd
	set omd = New CateEventBanner
	omd.FCurrPage = page
	omd.FPageSize=10
	omd.FRectDisp = vCateCode
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
		frm.action="doMainEventBanner2015.asp";
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
		upfrm.action="doMainEventBanner2015.asp";
		upfrm.submit();

	}
}

// �̺�Ʈ ��� ������ ���뿩�� Ȯ��
function RefreshCategoryEventBanner(){
    if (confirm('�����Ͻðڽ��ϱ�?')){
		 refreshFrm.target = "prociframe";
		 refreshFrm.action = "https://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/chtml/dispcate/catemain_eventbanner_make.asp";
		 refreshFrm.submit();
    }
}

function RefreshCategoryEventBannerTest(){
    if (confirm('�����Ͻðڽ��ϱ�?')){
		 refreshFrm.target = "prociframe";
		 refreshFrm.action = "https://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/chtml_test/dispcate/catemain_eventbanner_make.asp";
		 refreshFrm.submit();
    }
}


// ���� ����
function viewPage(idx)
{
	frm.mode.value="edit";
	frm.page.value=<%=page%>;
	frm.idx.value=idx;
	frm.action="category_main_EventBanner_input2015.asp";
	frm.submit();
}

</script>
<br />
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="post">
<input type="hidden" name="gb" value="proc">
<input type="hidden" name="catecode" value="<%=vCateCode%>">
</form>
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
			����ī�װ� :
			<%
			Dim cDisp
			SET cDisp = New cDispCate
			cDisp.FCurrPage = 1
			cDisp.FPageSize = 2000
			cDisp.FRectDepth = 1
			'cDisp.FRectUseYN = "Y"
			cDisp.GetDispCateList()

			If cDisp.FResultCount > 0 Then
				Response.Write "<select name=""catecode"" class=""select"" onChange=""frm.submit();"">" & vbCrLf
				Response.Write "<option value="""">����</option>" & vbCrLf
				For i=0 To cDisp.FResultCount-1
					Response.Write "<option value=""" & cDisp.FItemList(i).FCateCode & """ " & CHKIIF(CStr(vCateCode)=CStr(cDisp.FItemList(i).FCateCode),"selected","") & ">" & cDisp.FItemList(i).FCateName & "</option>"
				Next
				Response.Write "</select>&nbsp;&nbsp;&nbsp;"
			End If
			Set cDisp = Nothing
			%>
			������� : <select name="isusing" onChange="frm.submit();"><option value="Y" <%=CHKIIF(isusing="Y","selected","")%>>Yes</option><option value="N" <%=CHKIIF(isusing="N","selected","")%>>No</option></select>
			&nbsp;&nbsp;&nbsp;
			�̺�Ʈ�ڵ� : <input type="text" name="evtCd" value="<%=evtCd%>" size="6">
		</td>
		<td>
			<input type="button" value="�� ��" onclick="frm.submit();">
		</td>
	</tr>
	<%IF vCateCode <> "" THEN%>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="50" colspan="2">
			<a href="javascript:RefreshCategoryEventBanner()"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle"><b>Real ����</b></a>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="javascript:RefreshCategoryEventBannerTest()"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle"><b>�׽�Ʈ ����</b></a>
			->
			<a href="https://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/shopping/category_main_test.asp?disp=<%=vCateCode%>" target="_blank"><b>[�׽�Ʈ ������ Ȯ���ϱ�]</b></a>
		</td>
	</tr>
	<%END IF%>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr height="50">
		<td align="left">
			�����Ѱ�
			<select name="allusing"><option value="Y">������� Y ����</option><option value="N">������� N ����</option></select> <input type="button" class="button" value="����" onclick="changeUsing(frm);">
		</td>
		<td align="right">
			<input type="button" value="������ �߰�" onclick="self.location='/admin/categorymaster/category_main_EventBanner_input2015.asp?mode=add&catecode=<%= vCateCode %>&menupos=<%= menupos %>'" class="button">
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
		<td width="50" align="center">���Ĺ�ȣ</td>
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
		<td align="center"><%= omd.FItemList(i).Fcode_nm %></td>
		<td align="center">
			<a href="javascript:viewPage(<%= omd.FItemList(i).fidx %>);"><%= "[" & omd.FItemList(i).Fevt_code & "] " & omd.FItemList(i).Fevt_name %></a>
			<br />
			<%= omd.FItemList(i).Fevt_subcopykor %>
			<% If omd.FItemList(i).Fevt_stdt <> "" Then %>
			<br/><br/>
			�̺�Ʈ �Ⱓ : <span style="color:red"><%=omd.FItemList(i).Fevt_stdt %>~<%=omd.FItemList(i).Fevt_etdt %></span>
			<% End If %>
		</td>
		<td align="center">
			<img src="<%= omd.FItemList(i).Fevt_molistbanner %>" width="100" border="0">
		</td>
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
				<a href="?page=<%= omd.StarScrollPage-1 %>&menupos=<%= menupos %>&isusing=<%=isusing%>&catecode=<%=vCateCode%>&evtCd=<%=evtCd%>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + omd.StarScrollPage to omd.FScrollCount + omd.StarScrollPage - 1 %>
				<% if i>omd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&catecode=<%=vCateCode%>&evtCd=<%=evtCd%>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if omd.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&catecode=<%=vCateCode%>&evtCd=<%=evtCd%>">[next]</a>
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
<iframe src="" name="prociframe" id="prociframe" width="0" height="0" frameborder="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->