<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/academy/lib/classes/sitemaster/main_TopKeywrdCls.asp"-->
<%
'###############################################
' PageName : ž�޴��˻�������
' Discription : ���� ž Ű���� ���
' History : 2009.09.16 �ѿ�� 10x10���� ������ ����
'###############################################

dim page, SearchString, strUse, lp , i ,oKeyword , keyword_gubun
	page = RequestCheckvar(request("page"),10)
	SearchString = request("SearchString")
	keyword_gubun = RequestCheckvar(request("keyword_gubun"),10)
	strUse = RequestCheckvar(request("strUse"),1)
	if page = "" then page=1
	if strUse = "" then strUse="Y"

set oKeyword = New CSearchKeyWord
	oKeyword.FCurrPage = page
	oKeyword.FPageSize=20
	oKeyword.frectkeyword_gubun = keyword_gubun
	oKeyword.FRectUsing = strUse
	oKeyword.FRectSearch = SearchString
	oKeyword.GetSearchKeyWord()
%>
<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	var frm = document.frmBuyPrc;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			frm.selIdx[i].checked = bool;
		}
	} else {
		frm.selIdx.checked = bool;
	}
}

function CheckSelected(){
	var pass = false;
	var frm = document.frmBuyPrc;

	if(frm.selIdx.length) {
		for (var i=0;i<frm.selIdx.length;i++){
			pass = ((pass)||(frm.selIdx[i].checked));
		}
	} else {
		pass = ((pass)||(frm.selIdx.checked));
	}

	if (!pass) {
		return false;
	}
	return true;
}

function changeUsing(upfrm){
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}
	
	if (upfrm.allusing.value=='Y'){
		var ret = confirm('���� �������� ��������� �����մϴ�');
	} else {
		var ret = confirm('���� �������� ���������� �����մϴ�');
	}

	if (ret) {
		upfrm.mode.value="changeUsing";
		upfrm.action="doMainTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}


function changeSort(upfrm){
	var arrSort="";
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if(confirm('���� �����ۿ� �Է��Ͻ� ������ȣ��� �����մϴ�.')) {

		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked) arrSort = arrSort + upfrm.SortNo[i].value + ",";
			}
		} else {
			if(upfrm.selIdx.checked) arrSort=upfrm.SortNo.value;
		}
		upfrm.arrSort.value = arrSort;

		upfrm.mode.value="changeSort";
		upfrm.action="doMainTopKeyword.asp";
		upfrm.submit();
	} else {
		return;
	}
}

function frm_search()
{
	refreshFrm.target = "";
	refreshFrm.action = "main_TopKeyword.asp";
}

	// ������ �̵�
	function goPage(pg)
	{
		document.refreshFrm.page.value=pg;
		document.refreshFrm.action="main_TopKeyword.asp";
		document.refreshFrm.submit();
	}

// ī�װ� ����� ���
function changecontent() {
}

function AssignReal(upfrm , keyword_gubun, device){
	var idxarr; 
	var tmp;
	tmp =0;
	idxarr = "";
	
	if (!CheckSelected()){
		alert('���þ������� �����ϴ�.');
		return;
	}

	if(confirm('�����Ͻðڽ��ϱ�?')) {
		if(upfrm.selIdx.length) {
			for (var i=0;i<upfrm.selIdx.length;i++){
				if(upfrm.selIdx[i].checked){
					idxarr = idxarr + upfrm.selIdx[i].value + ",";
					tmp = tmp + 1
				}	
			}
		}
	}else{
		return;
	}
	
	if (keyword_gubun == '0'){
		if (tmp > 3){
			alert('����[����]�� �˻��� ���� 3�������� �����մϴ�.');
			return;
		}	
	}else if(keyword_gubun == '1'){
		if (tmp > 7){
			alert('����[�˻�]�� �˻��� ���� 7�������� �����մϴ�.');
			return;
		}	
	}else if(keyword_gubun == '3'){
		if (tmp > 1){
			alert('��� �ؽ�Ʈ�� 1�������� �����մϴ�.');
			return;
		}
		idxarr = upfrm.selIdx.value+",";
	}
	if(device == "W"){
		AssignbestReal = window.open("<%=www1Fingers%>/chtml/make_keyword.asp?idxarr=" +idxarr+ "&keyword_gubun="+keyword_gubun, "AssignbestReal","width=400,height=300,scrollbars=yes,resizable=yes");
	}else{
		AssignbestReal = window.open("<%=mob1Fingers%>/chtml/make_keyword.asp?idxarr=" +idxarr+ "&keyword_gubun="+keyword_gubun, "AssignbestReal","width=400,height=300,scrollbars=yes,resizable=yes");
	}
	AssignbestReal.focus();
}

</script>

<!-- ��� �˻��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="refreshFrm" method="get" onSubmit="frm_search()" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td>
		<table width="100%" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				���� : <% drawkeyword_gubun "keyword_gubun",keyword_gubun %>
				��뿩��
				<select class="select" name="strUse">
					<option value="all">��ü</option>
					<option value="Y">���</option>
					<option value="N">����</option>
				</select>
				/ Ű���� �˻�
				<input type="text" class="text" name="SearchString" size="12" value="<%=SearchString%>">
				<script language="javascript">
					document.refreshFrm.strUse.value="<%=strUse%>";
				</script>
			</td>
		</tr>
		</table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<form name="frmBuyPrc" method="post" action="">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="arrSort" value="">
<tr>
	<td>
		<% if keyword_gubun <> "" then %>
			<input type="button" onclick="AssignReal(frmBuyPrc,'<%=keyword_gubun%>','W')" value="PC�Ǽ�������" class="button">
			&nbsp;<input type="button" onclick="AssignReal(frmBuyPrc,'<%=keyword_gubun%>','M')" value="Mobile�Ǽ�������" class="button">
		<% end if %>
	</td>
	<td align="right">
		<select class="select" name="allusing">
			<option value="Y">���� -> Y</option>
			<option value="N">���� -> N</option>
		</select>
		<input type="button" class="button" value="����" onclick="changeUsing(frmBuyPrc);">
		/
		<input type="button" class="button" value="��������" onclick="changeSort(frmBuyPrc);">
		/
		<input type="button" value="������ �߰�" onclick="self.location='main_TopKeyword_write.asp?menupos=<%= menupos %>'" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=oKeyword.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oKeyword.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>��ȣ</td>
	<td>����</td>
	<td>Ű����</td>
	<td>��ũ����</td>
	<td>�������</td>
	<td>����</td>
	<td>�����</td>
</tr>
<%	if oKeyword.FResultCount < 1 then %>
<tr>
	<td colspan="8" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for i=0 to oKeyword.FResultCount-1
%>
<tr align="center" bgcolor="<% if oKeyword.FItemList(i).Fisusing = "Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><input type="checkbox" name="selIdx" value="<%= oKeyword.FItemList(i).Fidx %>"></td>
	<td><%= oKeyword.FItemList(i).Fidx %></td>
	<td><%= drawkeyword_gubunname(oKeyword.FItemList(i).fkeyword_gubun) %></td>
	<td><a href="main_TopKeyword_write.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Fkeyword %></a></td>
	<td align="left" style="word-break : break-all;">&nbsp;<a href="main_TopKeyword_write.asp?idx=<%= oKeyword.FItemList(i).Fidx %>&page=<%=page%>"><%= oKeyword.FItemList(i).Flinkinfo %></a></td>
	<td><%=oKeyword.FItemList(i).Fisusing%></td>
	<td><input type="text" class="text" name="SortNo" value="<%=oKeyword.FItemList(i).FsortNo%>" size="2" style="text-align:center"></td>
	<td><%= FormatDate(oKeyword.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<tr bgcolor="#FFFFFF">
	<td colspan="8" align="center">
	<!-- ������ ���� -->
	<%
		if oKeyword.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oKeyword.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oKeyword.StartScrollPage to oKeyword.FScrollCount + oKeyword.StartScrollPage - 1

			if lp>oKeyword.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oKeyword.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</form>
</table>

<%
set oKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
