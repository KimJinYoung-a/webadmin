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
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<%
'###############################################
' PageName : pcmain_manager.asp
' Discription : ����Ʈ ���� ����
' History : 2018-03-05 ����ȭ
'###############################################

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun, targetUser , prevTime
targetUser = "��ü"
dim page,strParm, datediv
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")
	datediv = request("datediv")
	prevTime = request("prevTime")

	If gubun = "" Then
		gubun = "index"
	End If

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if

	if prevTime = "" then prevTime = "00"

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
		if (oposcode.FOneItem.Ffixtype="D" Or poscode="714" Or poscode="710") then
		'���ں��϶� ������ �̸����� ��¥ ����
		oMainContents.FRectDateDiv = datediv

		end if
	oMainContents.Flinktype = oposcode.FOneItem.Flinktype
	oMainContents.FRectSelDate = prevDate
	oMainContents.FRectSelDateTime = prevTime
	end if
	oMainContents.GetMainContentsList

dim i


	'### ���к� js �������� ### (���� index, �ΰŽ�, ����Ʈ������ ���� ������̾ �״�� ���. ���� ���濹��.
	Dim vGubun
	If gubun = "my10x10" Then
		vGubun = "_my10x10"
	End IF
%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
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

function setDefault()
{
	frm.poscode.options[0].selected = true;
	frm.submit();
}
</script>

<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    ��뱸��
		<select name="isusing" class="select">
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

		<% If  poscode="714" Or poscode="710" Then %>
		<select name="datediv" class="select">
		<option value="1" <% if datediv="1" then response.write "selected" %> >������
		<option value="2" <% if datediv="2" then response.write "selected" %> >������
		</select>
		<input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<% Else %>
        &nbsp;&nbsp;
        �������� <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<% if prevDate <> "" then %>
		�ð� <input type="input" name="prevTime" value="<%=prevTime%>" class="text" size="2" maxlength="2" /> ��~
		<% end if %>
		<% End If %>
		<br>
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >��������
	    <br>
	    �� <font color="blue">�׷챸�� : index - 10x10 ����</font><br/>
	    �� <font color="blue">�׷챸�� : PCbanner - 10x10 PC ���</font><br/>
	    �� <font color="blue">�׷챸�� : MAbanner - 10x10 M/A ���</font>
		</font>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
<!--<td><a href="http://www.10x10.co.kr/index_preview.asp?yyyymmdd=<%= Left(CStr(now()),10) %>" target="refreshFrm_Main">�������</a></td>-->
    <td colspan="13" align="right">
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
    <td>����ī�װ�</td>	
    <td>��ũ<br>����</td>
    <td>�ݿ�<br>�ֱ�</td>
    <td>������</td>
    <td>������</td>
    <td>��뿩��</td>
	<td>������</td>
    <td>�켱����</td>
    <td>�����</td>
    <td>�۾���</td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1	
	
	if not isnull(oMainContents.FItemList(i).FtargetType) then
		Select Case cstr(oMainContents.FItemList(i).FtargetType)
			Case ""
				targetUser = "����"	
			Case "0"
				targetUser = "white"
			Case "1"
				targetUser = "red"			
			Case "2"
				targetUser = "vip"			
			Case "3"
				targetUser = "vip gold"			
			Case "4"
				targetUser = "vvip"
			Case "4"
				targetUser = "vvip"
			Case "7"
				targetUser = "STAFF"
			Case "8"
				targetUser = "FAMILY"
			Case "9"
				targetUser = "BIZ"
			case "00"
				targetUser = "ȸ����ü"
			case "99"
				targetUser = "��ȸ��"
		end select
	else
		targetUser = "����"
	end if 
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).Fidx & "</a>" %></td>
    <td align="center"><a href="?gubun=<%=gubun%>&poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td align="center">
	<%
		'�ؽ�Ʈ ��ũŸ���̸� �ؽ�Ʈ ǥ�� - �ƴϸ� ������� �̹���
		if oMainContents.FItemList(i).Flinktype="T" then
			Response.Write "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).FlinkText & "</a>"
		Else
			If oMainContents.FItemList(i).Fposcode = "714" Then
	%>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).Fcultureimage %>" border="0" width=160 height=238 alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
    <% ElseIf oMainContents.FItemList(i).Fposcode = "706" Then %>   
			(�̹��� <%=oMainContents.FItemList(i).Fbannertype%>��)&nbsp;&nbsp;<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
	<% Else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
		<% If oMainContents.FItemList(i).GetImageUrl2 <> "" Then %>
    		<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).GetImageUrl2 %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname2 %>"></a>
		<% End If %>
		<% If oMainContents.FItemList(i).GetImageUrl3 <> "" Then %>
    		<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).GetImageUrl3 %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname3 %>"></a>
		<% End If %>		
    <%
			End If
    	end if
    %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).getDispCateListName %></td>
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
	<td align="center"><%= targetUser %></td>
    <td align="center">
    	<%
			response.write oMainContents.FItemList(i).forderidx
    	%>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Fregname %></td>
    <td align="center"><%= oMainContents.FItemList(i).Fworkername %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center" height="30">
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
