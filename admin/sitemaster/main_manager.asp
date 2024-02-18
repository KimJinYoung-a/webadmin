<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2008.04.11 ������ : �Ǽ������� ����
'			2009.04.19 �ѿ�� 2009�� �°� ����
'           2009.12.21 ������ : ���ں� �÷��� ���� ��� �߰�
'           2011.04.20 ������ : Text��ũ �߰�
'           2012.02.15 ������ : �̴ϴ޷� ��ü
'###############################################

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun
dim page,strParm
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")

	If gubun = "" Then
		gubun = "index"
	End If

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if

	if page="" then page=1
strParm = "isusing="&isusing&"&poscode="&poscode&"&fixtype="&fixtype&"&validdate="&validdate&"&prevDate="&prevDate&"&gubun="&gubun
dim oposcode
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode

	if (poscode<>"") then
	    oposcode.GetOneContentsCode
	end if

dim oMainContents
	set oMainContents = new CMainContents
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectfixtype = fixtype
	oMainContents.FRectPosCode = poscode
	oMainContents.Fgubun = gubun
	oMainContents.FRectvaliddate = validdate
	if (poscode<>"") then
		if (oposcode.FOneItem.Ffixtype="D") then
		'���ں��϶� ������ �̸����� ��¥ ����
		oMainContents.FRectSelDate = prevDate
		end if
	oMainContents.Flinktype = oposcode.FOneItem.Flinktype
	end if
	oMainContents.GetMainContentsList

dim i


	'### ���к� js �������� ### (���� index, �ΰŽ�, ����Ʈ������ ���� ������̾ �״�� ���. ���� ���濹��.
	Dim vGubun
	If gubun = "my10x10" Then
		vGubun = "_my10x10"
	End IF
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/sitemaster/lib/popmainposcodeedit.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/lib/popmaincontentsedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AssignTest(){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main_Test','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main_Test";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_contents_Test_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}

function AssignReal(){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_contents_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}


function AssignDailyTest(idx){
	 var popwin = window.open('','refreshFrm_Main_Test','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main_Test";
	 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_contents_byidx_Test_JS.asp?idx=" + idx;
	 refreshFrm.submit();
}

function AssignDailyReal(idx){
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";

	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "http://2011www.10x10.co.kr/chtml/make_main_contents_byidx_JS<%=vGubun%>.asp?idx=" + idx;
	 <% else %>
	 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_contents_byidx_JS<%=vGubun%>.asp?idx=" + idx;
	 <% end if %>

	 refreshFrm.submit();
}


function AssignFlashReal(pc,lt){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";

		 if(lt=="F") {
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_flash_Text.asp?poscode=" + document.frm.poscode.value;
		 } else {
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_Button_Text.asp?poscode=" + document.frm.poscode.value;
		 }
			 refreshFrm.submit();
	}
}

function AssignFlashDailyReal(pc,lt,vt){
    if (document.frm.poscode.value == ""){
		alert("������ġ�� �������ּ���");
		document.frm.poscode.focus();
	} else if (vt<=0 || !vt){
		alert("���� �Ⱓ�� �Է����ּ���.");
	}else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";

		 if(lt=="F") {
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_flash_JS.asp?poscode=" + document.frm.poscode.value + "&vTerm=" + vt;
		 } else {
			 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_Button_JS.asp?poscode=" + document.frm.poscode.value + "&vTerm=" + vt;
		 }
			 refreshFrm.submit();
	}
}

function AssignRealApplALL(allrefreshVal){
     if (!confirm('���� �ݿ��Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "<%=wwwUrl%>/chtml/make_mainApp_refresh.asp?allrefresh=" + allrefreshVal;
	 <% else %>
	 refreshFrm.action = "/admin/sitemaster/lib/popAppAssign.asp?allrefresh=" + allrefreshVal;
	 <% end if %>
	 refreshFrm.submit();
}

function AssignRealAppl(idx){
    if (!confirm('���� �ݿ��Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "<%=wwwUrl%>/chtml/make_mainApp_refresh.asp?idx=" + idx;
	 <% else %>
	 refreshFrm.action = "/admin/sitemaster/lib/popAppAssign.asp?idx=" + idx;
	 <% end if %>
	 refreshFrm.submit();
}

function AssignXmlAppl(term){
    if (!confirm('���� �ݿ��Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "http://2015www.10x10.co.kr/chtml/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term +"&rid=<%=session("ssBctId")%>";
	 <% else %>
	 refreshFrm.action = "http://www1.10x10.co.kr/chtml/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term +"&rid=<%=session("ssBctId")%>";
	 <% end if %>
	 refreshFrm.submit();

	 setTimeout("location.reload()",1000);
}

function setDefault()
{
	frm.poscode.options[0].selected = true;
	frm.submit();
}


//test XML ���� 2016-03-22 ���¿�
function AssignXmlAppltest(term){
    if (!confirm('TEST ������ �ݿ��Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "http://2015www.10x10.co.kr/chtml_test/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
	 <% else %>
	 refreshFrm.action = "http://www1.10x10.co.kr/chtml_test/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
	 <% end if %>
	 refreshFrm.submit();

	 //Staging������ ���� 2018-01-25 ������ >> Rollback ������
//	 var popwin = window.open('','refreshFrm_Main','');
//	 popwin.focus();
//	 refreshFrm.target = "refreshFrm_Main";
//	 refreshFrm.action = "http://stgwww.10x10.co.kr/chtml_test/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
//	 refreshFrm.submit();
}

//test �̸����� 2016-03-22 ���¿�
function testmainpage(){
	var yyyymmdd
		yyyymmdd = document.frmtest.prevDatetest.value;

	if(yyyymmdd==""){
		alert('�̸� �� ��¥�� ������ �ּ���.');
		document.frmtest.prevDatetest.focus();
		return;
	}
	var openNewWindow = window.open("about:blank");
	 <% if (application("Svr_Info")	= "Dev") then %>
	 	openNewWindow.location.href = "http://2015www.10x10.co.kr/_index_admin_test.asp?CtrltestDate="+yyyymmdd;
	 <% else %>
	 	openNewWindow.location.href = "http://www1.10x10.co.kr/_index_admin_test.asp?CtrltestDate="+yyyymmdd;
	 <% end if %>
	 return false;
}

</script>

<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">

	    ��뱸��
		<select name="isusing" class="select">
		<option value="">��ü
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >�����
		<option value="N" <% if isusing="N" then response.write "selected" %> >������
		</select>
		&nbsp;&nbsp;
		���뱸��
		<% call DrawFixTypeCombo ("fixtype", fixtype, "") %>

		&nbsp;&nbsp;
		�׷챸��
		<% call DrawGroupGubunCombo ("gubun", gubun, "onChange='setDefault()'") %>

		&nbsp;&nbsp;
		������ġ
		<% call DrawMainPosCodeCombo("poscode",poscode, "", gubun) %>
		<%
			if (poscode<>"") then
				if (oposcode.FOneItem.Ffixtype="D") then
		%>
        &nbsp;&nbsp;
        �������� <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<%
				end if
			end if
		%>

		<br>
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
	    <br>
	    �� <font color="blue">�׷챸�� : index - 10x10 ����</font>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td><a href="http://www.10x10.co.kr/index_preview.asp?yyyymmdd=<%= Left(CStr(now()),10) %>" target="refreshFrm_Main">�������</a></td>
    <td colspan="2">
	    <%
	    	if (poscode<>"") then
	    		if (oposcode.FOneItem.Ffixtype="R") AND gubun = "index" then
	    		'�ǽð� �ݿ�
	    %>
			        <a href="javascript:AssignRealApplALL('header');"><img src="/images/refreshcpage.gif" border="0"> Real ����(�ǽð� �ݿ��ֱ� �����ü)</a>
			        <a href="javascript:AssignRealApplALL('idx');"><img src="/images/refreshcpage.gif" border="0"> Real ����(�ǽð� �ݿ��ֱ� �ε�����ü)</a>
		<%
				elseif oposcode.FOneItem.Flinktype="F" or oposcode.FOneItem.Flinktype="B" then
					if (oposcode.FOneItem.Ffixtype="D") then
					'�÷��� ���ں� ����
		%>
						������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�
						<a href="javascript:AssignFlashDailyReal('<%= poscode %>','<%=oposcode.FOneItem.Flinktype%>',document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����(����)</a>
		<%
					else
					'�÷��� �Ϲ� ����
		%>
						<a href="javascript:AssignFlashReal('<%= poscode %>','<%=oposcode.FOneItem.Flinktype%>');"><img src="/images/refreshcpage.gif" border="0"> Flash Real ����</a>
		<%
					end if
				elseif (oposcode.FOneItem.Ffixtype <> "D") and (oposcode.FOneItem.Ffixtype <> "R") and (oposcode.FOneItem.Flinktype <> "X") and (oposcode.FOneItem.Flinktype <> "M") then
				'��ũ �� �Ϲ�
		%>
		    	    <!--<a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a>
		    	    &nbsp;&nbsp;//-->
		    	    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
	    <%
	    		elseif oposcode.FOneItem.Flinktype="X" or oposcode.FOneItem.Flinktype="M" Then
	    			if (oposcode.FOneItem.Ffixtype="D") then
		%>
						������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�
						<a href="javascript:AssignXmlAppl(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML(or ��) Real ����(����)</a>

						<form name="frmtest" method="get" action="">
						/////<b>TEST :</b>
						<a href="javascript:AssignXmlAppltest(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML(or ��) Staging ����(����)</a>
				        &nbsp;&nbsp;
				        �������� <input id="prevDatetest" name="prevDatetest" value="" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_triggertest" border="0" style="cursor:pointer" align="absmiddle" />
				        �� �̸�����-><input type="submit" class="button_s" onclick="testmainpage(); return false;" value="Ȯ��">
						</form>
						<script language="javascript">
							var CAL_Start = new Calendar({
								inputField : "prevDatetest", trigger    : "prevDate_triggertest",
								onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
							});
						</script>
		<%
					else
		%>
						<a href="javascript:AssignXmlAppl('');"><img src="/images/refreshcpage.gif" border="0"> XML(or ��) Real ����</a>
		<%
					end if
				end if
	    	end if

	    	If poscode <> "" Then
	    		Response.Write "&nbsp;&nbsp;&nbsp;" & fnMainManageOpenLog(poscode)
	    	End IF
	    %>
    </td>
    <td colspan="10" align="right">
    	<% if C_ADMIN_AUTH then %>
		<input type="button" class="button" value="�ڵ����" onClick="popPosCodeManage();">&nbsp;
		<% end if %>
    	<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>idx</td>
    <td>���и�</td>
    <td>�̹���/�ؽ�Ʈ</td>
    <td>��ũ<br>����</td>
    <td>�ݿ�<br>�ֱ�</td>
    <td>������</td>
    <td>������</td>
    <td>��뿩��</td>
    <td>�켱����</td>
    <td>�����</td>
    <td>�۾���</td>
    <td></td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).Fidx & "</a>" %></td>
    <td align="center"><a href="?gubun=<%=gubun%>&poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td>
	<%
		'�ؽ�Ʈ ��ũŸ���̸� �ؽ�Ʈ ǥ�� - �ƴϸ� ������� �̹���
		if oMainContents.FItemList(i).Flinktype="T" then
			Response.Write "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).FlinkText & "</a>"
		Else
			'�̹��� ����� ���� ǥ��(���� 300px)
			if oMainContents.FItemList(i).Fimagewidth>300 and Not(oMainContents.FItemList(i).getImageUrl="" or isNull(oMainContents.FItemList(i).getImageUrl)) then
	%>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=300 alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
	<%		else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" <% if InStr(gubun,"banner") > 0 Then  %>width=600<% End If %> alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
    <%
    		end if
    	end if
    %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).getlinktypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).getfixtypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).FStartdate %></td>
    <td align="center">
    <% if (oMainContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oMainContents.FItemList(i).FEnddate,10) %></font>
    <% else %>
    <%= Left(oMainContents.FItemList(i).FEnddate,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
    <td align="center">
    	<%
    	'// ������ ������ġ���� �켱���� ���
    	'Select Case poscode
    	'	Case "400", "401", "402", "403", "404", "405", "420", "421", "428"
    			response.write oMainContents.FItemList(i).forderidx
    	'end Select
    	%>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Fregname %></td>
    <td align="center"><%= oMainContents.FItemList(i).Fworkername %></td>
    <td>
    <% if (oMainContents.FItemList(i).Ffixtype="R") then %>
	<% if InStr(gubun,"banner") < 0 Then  %>
    <a href="javascript:AssignRealAppl('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
	<% End If %>
    <% elseif Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype="F" or oMainContents.FItemList(i).Flinktype="B" or oMainContents.FItemList(i).Ffixtype="R") then %>
    <!--<a href="javascript:AssignDailyTest('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/icon_search.jpg" border="0"> �̸�����</a> //-->
    	<% If oMainContents.FItemList(i).Flinktype <> "X" and oMainContents.FItemList(i).Flinktype <> "M" Then %>
    		&nbsp;
    		<a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
    	<% Else %>

    	<% End If %>
    <% else %>
    &nbsp;
    <% end if %>
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="12" align="center" height="30">
    <% if oMainContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oMainContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oMainContents.StarScrollPage to oMainContents.FScrollCount + oMainContents.StarScrollPage - 1 %>
		<% if i>oMainContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oMainContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->