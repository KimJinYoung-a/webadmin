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
<!-- #include virtual="/admin/mobile/submenu/inc_subhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/TopcateManageCls.asp" -->
<!-- #include virtual="/lib/classes/mobile/topcatecodeCls.asp" -->
<%
'###############################################
' PageName : main_manager.asp
' Discription : ����� ����Ʈ ���� ����
' History : 2010.02.23 ������
'           2011.12.23 ������ : ���ں� ���� ��� �߰�
'           2012.02.14 ������ : �̴ϴ޷� ��ü
'###############################################

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate , sedatechk
dim page , gcode
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gcode = request("gcode")

	sedatechk = request("sedatechk")

	if ((research="") and (isusing="")) then 
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

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
	oMainContents.FRectvaliddate = validdate
	oMainContents.FRectSelDate = prevDate
	oMainContents.FRectgnbcode = gcode
	oMainContents.FRectsedatechk= sedatechk '//������ ���� üũ
	if (poscode<>"") then
	oMainContents.Flinktype = oposcode.FOneItem.Flinktype
	end if
	oMainContents.GetMainContentsList

dim i
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
    var popwin = window.open('popMainPoscodeEdit.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx,gcode,poscode){
    var popwin = window.open('popmaincontentsedit.asp?idx=' + idx +'&gcode='+ gcode +'&poscode='+ poscode,'mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//function AssignReal(){
//    if (document.frm.poscode.value == ""){
//		alert("������ġ�� �������ּ���");
//		document.frm.poscode.focus();
//	}
//	else{
//		 var popwin = window.open('','refreshFrm_Main','');
//		 popwin.focus();
//		 refreshFrm.target = "refreshFrm_Main";
//		 refreshFrm.action = "<%=wwwUrl%>/chtml/mobile/make_main_contents_JS.asp?poscode=" + document.frm.poscode.value;
//		 refreshFrm.submit();
//	}
//}
//
//
//function AssignDailyReal(idx){
//	 var popwin = window.open('','refreshFrm_Main','');
//	 popwin.focus();
//	 refreshFrm.target = "refreshFrm_Main";
//	 refreshFrm.action = "<%=wwwUrl%>/chtml/mobile/make_main_contents_byidx_JS.asp?idx=" + idx;
//	 refreshFrm.submit();
//}

function AssignXMLReal(term){
	if (!confirm('���� �ݿ��Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_cate_main_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term + "&gnbcode=" + document.frm.gcode.value;
	 refreshFrm.submit();
}
//function AssignJSReal(){
//    if (document.frm.poscode.value == ""){
//		alert("������ġ�� �������ּ���");
//		document.frm.poscode.focus();
//	}
//	else{
//		 var popwin = window.open('','refreshFrm_Main','');
//		 popwin.focus();
//		 refreshFrm.target = "refreshFrm_Main";
//		 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_main_poscode_JS.asp?poscode=" + document.frm.poscode.value;
//		 refreshFrm.submit();
//	}
//}

function DeleteReal(term){
	if (!confirm('���� �������� ��ʸ� ����(��������) �Ͻðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/delete_main_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
	 refreshFrm.submit();
}

function St_rotate(mna){
	if (!confirm('���� �Ѹ���� ������ �����ù�� �켱������ �ٲٽðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/make_main_rollchk.asp?mna="+mna;
	 refreshFrm.submit();
}

function Ed_rotate(mna){
	if (!confirm('���� �Ѹ���� ������ 1����� �켱������ �ٲٽðڽ��ϱ�?')) return;

	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 refreshFrm.action = "<%=mobileUrl%>/chtml/mobile/delete_main_rollchk.asp?mna="+mna;
	 refreshFrm.submit();
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
		* ���� ���� : <span style="font-size:13px;"><strong>GNB �޴� �˻��� XML ��� ��ư�� ���� �˴ϴ�. (���� ����)</strong></span></br>
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
	    &nbsp;
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
		������ġ
		<% call DrawMainPosCodeCombo("poscode",poscode, "") %>
        &nbsp;&nbsp;
		GNB ����
		<% Call drawSelectBoxGNB("gcode" , gcode) %>
        &nbsp;&nbsp;
		�����ϱ��� <input type="checkbox" name="sedatechk" <% if sedatechk="on" then response.write "checked" %> />
        �������� <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script type="text/javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
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
    <td>
	    <%
	    	if (poscode<>"" And gcode <> "") then
				if oposcode.FOneItem.Flinktype="X" then
				'XML
	    			if (oposcode.FOneItem.Ffixtype="D") then
		%>
						������ �����Ͽ� <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">�ϰ�
						<a href="javascript:AssignXMLReal(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML Real ����(����)</a>
		<%
					else
		%>
						<a href="javascript:AssignXMLReal('');"><img src="/images/refreshcpage.gif" border="0"> XML Real ����</a>
		<%
					end if
				elseif oposcode.FOneItem.Flinktype="J" then
		%>
						<a href="javascript:AssignJSReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> JS Real ����</a>
		<%
				elseif (oposcode.FOneItem.Ffixtype <> "D") then
				'��ũ �� �Ϲ�
		%>
    	    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
	    <%
	    		end if
	    	end if
	    %>
		<% If poscode = "2049" Or poscode = "2042" Or poscode = "2053" Or poscode = "2054" then	%>
			<a href="javascript:DeleteReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> XML ���� (���� ����)</a>
			<a href="javascript:St_rotate('<%=chkiif(poscode="2042" Or poscode="2053","m","a")%>');"><img src="/images/refreshcpage.gif" border="0"> �����ù�� �켱 �Ѹ�</a>
			<a href="javascript:Ed_rotate('<%=chkiif(poscode="2042" Or poscode="2053","m","a")%>');"><img src="/images/refreshcpage.gif" border="0"> 1����� �켱 �Ѹ�</a>
		<% End If %>
    </td>
    <td align="right">
    	<% if C_ADMIN_AUTH then %>
		<input type="button" class="button" value="�ڵ����" onClick="popPosCodeManage();">&nbsp;
		<% end if %>
    	<a href="javascript:AddNewMainContents('0','<%=gcode%>','<%=poscode%>');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		�˻���� : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="5%">idx</td>
    <td width="15%">���и�</td>
    <td width="*">�̹���</td>
    <td width="5%">��ũ<br>����</td>
    <td width="5%">�ݿ�<br>�ֱ�</td>
    <td width="5%">������</td>
    <td width="5%">������</td>
    <td width="5%">��뿩��</td>
    <td width="5%">�켱����</td>
    <td width="10%">�����</td>
    <td width="*"></td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= oMainContents.FItemList(i).Fidx %></td>
    <td align="left" style="padding-left:10px;"><a href="?poscode=<%= oMainContents.FItemList(i).Fposcode %>">������ġ : <%= oMainContents.FItemList(i).Fposname %></a><br><br>GNB�޴� : <%= oMainContents.FItemList(i).Fgnbname %></td>
    <td align="center">
	<%
		'�̹��� ����� ���� ǥ��(���� 300px)
		if oMainContents.FItemList(i).Fimagewidth>300 then
	%>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>','','');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=200 alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
	<% else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>','','');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=200 alt="<%=oMainContents.FItemList(i).Faltname%>"></a>
    <% end if %>
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
		response.write oMainContents.FItemList(i).forderidx 
    	%>
    </td>
    <td align="center"><%=getStaffUserName(oMainContents.FItemList(i).Freguserid)%></td>
    <td>
    <% if Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype="X" or oMainContents.FItemList(i).Flinktype="J") then %>
    <a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real ����</a>
    <% else %>
    &nbsp;
    <% end if %> 
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="center">
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