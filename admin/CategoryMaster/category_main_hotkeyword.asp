<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_hotkeywordcls.asp" -->
<%

'// ���� ����
dim cdl, cdm, page, isusing, divCd, vCatecode
vCateCode = Request("catecode")
page = request("page")
isusing = request("isusing")

if page="" then page=1
if isusing="" then isusing="Y"

dim omd
set omd = New CateHotKeyword
omd.FCurrPage = page
omd.FPageSize=8
omd.FRectIsusing = isusing
omd.FDisp = vCateCode
omd.GetPageItemList

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
		alert('���þ������� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �������� �����Ͻðڽ��ϱ�?');

	if (ret){
		upfrm.idx.value = "";
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + "," ;
				}
			}
		}
		upfrm.mode.value="del";
		upfrm.submit();

	}
}

// ��ü ������� ����
function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if (frm.allusing.value=='Y'){
		var ret = confirm('���� �������� ��������� �����մϴ�');
	} else {
		var ret = confirm('���� �������� ������ ����  �����մϴ�');
	}
	
	if (ret){
		upfrm.idx.value = "";
		var frm11;
		for (var i=0;i<document.forms.length;i++){
			frm11 = document.forms[i];
			if (frm11.name.substr(0,9)=="frmBuyPrc") {
				if (frm11.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm11.idx.value + "," ;
				}
			}
		}
		
	upfrm.isusing.value = frm.allusing.value;
	upfrm.mode.value="changeUsing";
	upfrm.submit();
	}
}

function popMainCodeManage(){
    var popwin = window.open('/admin/categorymaster/popMainPageCodeEdit.asp','popMainCode','width=800,height=600,scrollbars=yes');
    popwin.focus();
}

function AssignTest(){
    if (document.frm.divCd.value == ""){
		alert("�׸񱸺��� �������ּ���");
		document.frm.divCd.focus();
	}
	else if (document.frm.cdl.value == ""){
		alert("ī�װ��� �������ּ���");
		document.frm.cdl.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Cate','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Cate";
		 refreshFrm.action = "<%=uploadUrl%>/chtml/make_category_todayhot_test.asp?divCd=" + document.frm.divCd.value + "&cdl=" + document.frm.cdl.value;
		 refreshFrm.submit();
	}
}

function AssignReal(disp){
	if(confirm("�����Ͻðڽ��ϱ�?") == true) {
		 var todayhot = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/chtml/dispcate/catemain_hotkeyword_make.asp?catecode='+disp+'','todayhot','');
		 todayhot.focus();
	}
}

function AssignRealTest(disp){
	if(confirm("�����Ͻðڽ��ϱ�?") == true) {
		 var todayhot = window.open('http://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/chtml_test/dispcate/catemain_hotkeyword_make.asp?catecode='+disp+'','todayhot','');
		 todayhot.focus();
	}
}

function changecontent(){
	document.frm.submit();
}
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr>
		<td>
			�� �� ī�װ��� ���ο� ǥ�õǴ� ��Ű���� ��� ������ �Դϴ�.<br>
			�� ��������� Y �ΰ͵� �� ���Ĺ�ȣ�� ������(0�� ù��°) 4���� ǥ�õ˴ϴ�.<br>
			�� ��ǰ�ڵ�� �̹����� �ҷ����� ���� �Դϴ�. ���ḵũ�� ���� �־��ּž� �մϴ�.<br>
			�� ī�װ��� �����Ͻø� real ���� ��ư�� ���Դϴ�.<br>
		</td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="refreshFrm" method="post"></form>
<form name="frm" method="get" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="idxarr" value="">

<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="left">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
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
					/
					������� : <select name="isusing" onChange="frm.submit();"><option value="Y">Yes</option><option value="N">No</option></select>
					<script language="javascript">
						document.frm.isusing.value="<%=isusing%>";
					</script>
				</td>
			</tr>
			<tr>
				<td align="right">
				<select name="allusing"><option value="Y">���� -> Y</option><option value="N">���� ->N </option></select>
				<input type="button" class="button" value="����" onclick="changeUsing(delform);">
				<% if C_ADMIN_AUTH then %>
				<input type="button" value="�ڵ����" onClick="popMainCodeManage();" class="button">
				<% end if %>
				<input type="button" value="���þ����ۻ���" onclick="delitems(delform);" class="button">
				<input type="button" value="������ �߰�" onclick="self.location='/admin/categorymaster/category_main_hotkeyword_input.asp?mode=add&catecode=<%= vCateCode %>&divCd=<%= divCd %>&menupos=<%= menupos %>'" class="button">
			</td>
			</tr>
		</table>
	</td>
	<td valign="top" align="right">
		<!--<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">//-->
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- ��� �˻��� �� -->

<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<%IF vCateCode <> "" THEN%>
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="10" align="left" height="50"> 
		<a href="javascript:AssignReal('<%= vCateCode %>');"><img src="/images/refreshcpage.gif" border="0"><b> Real ����</b></a>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="javascript:AssignRealTest('<%= vCateCode %>')"><img src="/images/refreshcpage.gif" width="19" height="23" border="0" align="absmiddle"><b>�׽�Ʈ ����</b></a>
		->
		<a href="http://<%=CHKIIF(application("Svr_Info")="Dev","2015www","www1")%>.10x10.co.kr/shopping/category_main_test.asp?disp=<%=vCateCode%>" target="_blank"><b>[�׽�Ʈ ������ Ȯ���ϱ�]</b></a>
		</td>		
	</tr>
	<%END IF%>
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="10" align="left">�˻��Ǽ� : <%= omd.FTotalCount %> �� Page : <%= page %>/<%= omd.FTotalPage %></td>
		
	</tr>
	
	<tr align="center" bgcolor="#DDDDFF">
	<td width="50" align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td width="100" align="center">ī�װ���</td>
	<td width="154" align="center">��ǰ�ڵ�</td>
	<td width="154" align="center">��ǰ��</td>
	<td width="154" align="center">�̹���</td>
	<td align="center" width="154">Ű���幮��</td>
	<td align="center"width="154">��ũURL</td>
	<td width="50" align="center">����</td>
	<td width="50" align="center">�������</td>
	<td width="80" align="center">�����</td>
	</tr>
<% for i=0 to omd.FResultCount-1 %>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name="idx" value="<%= omd.FItemList(i).Fidx %>">
<tr bgcolor="#FFFFFF">
	<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td align="center"><%= omd.FItemList(i).Fcode_nm %></td>
	<td align="center"><%= omd.FItemList(i).Fitemid %></td>
	<td align="center"><%= omd.FItemList(i).Fitemname %></td>
	<td align="center"><img src="<%= omd.FItemList(i).FimgFile %>" width="150" border="0"></td>
	<td align="center">
		<a href="category_main_hotkeyword_input.asp?idx=<%= omd.FItemList(i).Fidx %>&mode=edit&menupos=<%=menupos%>"><%= omd.FItemList(i).Fkeyword %></a>
	</td>
	<td align="center"><a href="http://www.10x10.co.kr<%= omd.FItemList(i).FlinkURL %>" target="_blank">http://www.10x10.co.kr<%= omd.FItemList(i).FlinkURL %></a></td>
	<td align="center"><%= omd.FItemList(i).FSortNo %></td>
	<td align="center"><%= omd.FItemList(i).Fisusing %></td>
	<td align="center"><%= FormatDateTime(omd.FItemList(i).Fregdate,2) %></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">
	<% if omd.HasPreScroll then %>
		<a href="?page=<%= omd.StartScrollPage-1 %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&cdm=<%=cdm%>&catecode=<%=vCateCode%>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omd.StartScrollPage to omd.FScrollCount + omd.StartScrollPage - 1 %>
		<% if i>omd.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&cdm=<%=cdm%>&catecode=<%=vCateCode%>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omd.HasNextScroll then %>
		<a href="?page=<%= i %>&menupos=<%= menupos %>&isusing=<%=isusing%>&cdl=<%=cdl%>&cdm=<%=cdm%>&catecode=<%=vCateCode%>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr height="24" valign="bottom">
	<td><input type="button" value="���þ����ۻ���" onclick="delitems(delform);" class="button"></td>
	<td align="right">
		<% if C_ADMIN_AUTH then %>
		<input type="button" value="�ڵ����" onClick="popMainCodeManage();" class="button">
		<% end if %>
		<input type="button" value="������ �߰�" onclick="self.location='/admin/categorymaster/category_main_hotkeyword_input.asp?mode=add&catecode=<%= vCateCode %>&menupos=<%= menupos %>'" class="button">
	</td>
</tr>
</table>
<form name="delform" method="post" action="<%=uploadUrl%>/linkweb/doCategoryhotKeyword.asp" enctype="multipart/form-data">
<input type="hidden" name="catecode" value="<%= vCateCode %>">
<input type="hidden" name="mode">
<input type="hidden" name="idx">
<input type="hidden" name="isusing">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
</form>
<%
set omd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
