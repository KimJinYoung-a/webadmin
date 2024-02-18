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
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateMenuCls.asp"-->

<%
	Dim cMenu, vArr, i, vDisp1, vType, vPage, vUseYN, vOrderBy
	vPage = Request("page")
	vDisp1 = Request("disp1")
	vType = Request("type")
	vUseYN = Request("useyn")
	vOrderBy = Request("orderby")
	
	If vPage = "" Then vPage = "1" End If
	If vUseYN = "" Then vUseYN = "y" End If
	If vOrderBy = "" Then vOrderBy = "sortno asc" End If
	
	
	Set cMenu = New cDispCateMenu
	vArr = cMenu.GetDispCate1Depth()
	Set cMenu = Nothing
	
	Set cMenu = New cDispCateMenu
	cMenu.FCurrPage = vPage
	cMenu.FDisp1 = vDisp1
	cMenu.FType = vType
	cMenu.FUseYN = vUseYN
	cMenu.FOrderBy = vOrderBy
	cMenu.GetCateMainIssueList
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popreg(idx){
    var poprreg = window.open('reg.asp?disp1=<%=vDisp1%>&idx='+idx+'','poprreg','width=570,height=400');
    poprreg.focus();
}
function Check_All()
{
	var chk = f.idx;
	alert(chk.length);
	var cnt = 0;
	var ischecked = ""
	if(document.getElementById("chkall").checked){
		ischecked = "checked"
	}else{
		ischecked = ""
	}
	
	if(cnt == 0 && chk.length != 0){
		for(i = 0; i < chk.length; i++){ chk.item(i).checked = ischecked; }
		cnt++;
	}
}
function jsUseYNNO(){
	if(confirm("������ �͵��� �����Ͻðڽ��ϱ�?") == true) {
		f.action = "proc.asp";
		f.submit();
	}
}
function jsRealServerReg(){
	if(confirm("[<%=vDisp1%>] ī�װ��� �޴��� �����Ͻðڽ��ϱ�?") == true){
	    var popCreateTemp = window.open("http://<%=CHKIIF(application("Svr_Info")="Dev","2013www","www1")%>.10x10.co.kr/chtml/dispcate/menu_make_xml_New.asp?catecode=<%=vDisp1%>","popCreateTemp","width=1200 height=930 scrollbars=yes resizable=yes");
		popCreateTemp.focus();
	}
}
</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="40">
	<td align="left">
		&nbsp;
		<select name="disp1" class="select" onChange="frm.submit();">
		<option value="">-ī�װ�������-</option>
		<%
			For i=0 To UBound(vArr,2)
				Response.Write "<option value='" & vArr(0,i) & "' " & CHKIIF(CStr(vDisp1)=CStr(vArr(0,i)),"selected","") & ">" & vArr(1,i) & "</option>" & vbCrLf
			Next
		%>
		</select>
		&nbsp;&nbsp;&nbsp;
		<select name="type" class="select" onChange="frm.submit();">
			<option value="">-���м���-</option>
			<option value="issue_text" <%=CHKIIF(vType="issue_text","selected","")%>>issue_text</option>
			<option value="issue_image" <%=CHKIIF(vType="issue_image","selected","")%>>issue_image</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		<select name="useyn" class="select" onChange="frm.submit();">
			<option value="">-��뿩�μ���-</option>
			<option value="y" <%=CHKIIF(vUseYN="y","selected","")%>>���</option>
			<option value="n" <%=CHKIIF(vUseYN="n","selected","")%>>������</option>
		</select>
		&nbsp;&nbsp;&nbsp;
		<select name="orderby" class="select" onChange="frm.submit();">
			<option value="sortno asc" <%=CHKIIF(vOrderBy="sortno asc","selected","")%>>���Ĺ�ȣ��</option>
			<option value="regdate desc" <%=CHKIIF(vOrderBy="regdate desc","selected","")%>>�ֱٵ�ϼ�</option>
		</select>
	</td>
</tr>
</table>
</form>
<br>
<% If vDisp1 <> "" Then %>
	<% If vDisp1 = "110" Then %>
		Cat & Dog �޴� ������ ��а� �����ʽ��ϴ�. ��������� ������ ��û �ֽø� �˴ϴ�.
	<% Else %>
	<input type="button" value="[<%=vDisp1%>]ī�װ��� �޴� �����ϱ�" onClick="jsRealServerReg();">
	<% End If %>
<br>
<% End If %>
<br>
<form name="f" method="post" target="ifram">
<input type="hidden" name="action" value="del">
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="40" bgcolor="FFFFFF">
	<td colspan="12">
		<table width="100%" class="a">
		<tr>
			<td>
				�˻���� : <b><%=cMenu.FTotalCount%></b>
				&nbsp;
				������ : <b><%= vPage %> / <%=cMenu.FTotalPage%></b>
				<br>
				<input type="button" value="üũ�Ѱ� ������ ó��" onClick="jsUseYNNO();">
			</td>
			<td align="right">
				<input type="button" value="�űԵ��" onClick="popreg('');">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td><input type="checkbox" name="chkall" id="chkall" value="" onClick="Check_All()"></td>
    <td>idx</td>
    <td>����</td>
    <td>�̹���</td>
    <td>�ؽ�Ʈ</td>
    <td>������</td>
    <td>������</td>
    <td>��뿩��</td>
    <td>���Ĺ�ȣ</td>
    <td>�����</td>
    <td>�����</td>
    <td></td>
</tr>
<% If vDisp1 = "" Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" height="50">
		<center><b>* ī�װ����� ������ �ּ���.</b></center>
	</td>
</tr>
<% End If %>
<%
	for i=0 to cMenu.FResultCount - 1
%>
<% if cMenu.FItemList(i).Fuseyn = "n" then %>
<tr height="30" bgcolor="#DDDDDD">
<% else %>
<tr height="30" bgcolor="#FFFFFF">
<% end if %>
	<td align="center"><input type="checkbox" name="idx" value="<%= cMenu.FItemList(i).FIdx %>"></td>
	<td align="center"><%= cMenu.FItemList(i).FIdx %></td>
	<td align="center"><%= cMenu.FItemList(i).Ftype %></td>
	<td align="center"><% If cMenu.FItemList(i).Ftype = "issue_image" Then %><img src="<%= cMenu.FItemList(i).Fimgurl %>" height="29"><% End If %></td>
	<td align="center"><%= cMenu.FItemList(i).Fsubject %></td>
	<td align="center"><%= cMenu.FItemList(i).Fstartdate %></td>
	<td align="center"><%= cMenu.FItemList(i).Fenddate %></td>
	<td align="center"><%= cMenu.FItemList(i).Fuseyn %></td>
	<td align="center"><%= cMenu.FItemList(i).Fsortno %></td>
	<td align="center"><%= cMenu.FItemList(i).Fregusername %></td>
	<td align="center"><%= cMenu.FItemList(i).Fregdate %></td>
	<td align="center"><input type="button" value="����" onClick="popreg('<%= cMenu.FItemList(i).FIdx %>');"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" height="30">
    <td colspan="12" align="center">
    <% if cMenu.HasPreScroll then %>
		<a href="javascript:NextPage('<%= cMenu.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + cMenu.StartScrollPage to cMenu.FScrollCount + cMenu.StartScrollPage - 1 %>
		<% if i>cMenu.FTotalpage then Exit for %>
		<% if CStr(vPage)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if cMenu.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>

<iframe src="" name="ifram" width="0" height="0"></iframe>
<%
	Set cMenu = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->