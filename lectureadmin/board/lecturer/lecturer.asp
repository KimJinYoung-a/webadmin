<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ΰŽ� ���� �Խ���
' History : 2010.03.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/board/lecturer/lecturer_cls.asp"-->

<%
Dim i, vParam, sDoc_Status, sDoc_AnsOX, sSearchMine,sDoc_Type , page , g_MenuPos ,searchKey
dim searchString , Statusgubun
	Statusgubun = requestCheckVar(Request("Statusgubun"),10)
	searchKey		= requestCheckVar(Request("searchKey"),24)
	searchString		= requestCheckVar(Request("searchString"),32)	
	sDoc_Status		= requestCheckVar(Request("K000"),10)
	sDoc_Type		= NullFillWith(requestCheckVar(Request("G000"),10),"")
	sDoc_AnsOX		= NullFillWith(requestCheckVar(Request("ans_ox"),1),"")
	sSearchMine		= NullFillWith(requestCheckVar(Request("onlymine"),1),"o")
	g_MenuPos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	if page = "" then page = 1
	if sDoc_Status = "" and Statusgubun="" then 
		sDoc_Status = "K001"		
	end If
	If sDoc_Type="" Then
		sDoc_Type="G010"
	End If
	vParam = "K000="&sDoc_Status&"&s_type="&sDoc_Type&"&s_ans_ox="&sDoc_AnsOX&"&s_onlymine="&sSearchMine&"" &_
	+ "&searchKey="&searchKey&"&searchString="&searchString&"&Statusgubun="&Statusgubun&""

dim olect		
set olect = new clecturer_list
	olect.FPageSize = 20
	olect.FCurrPage = page
	olect.FrectDoc_Status = sDoc_Status
	olect.FrectDoc_Type = sDoc_Type
	olect.FrectDoc_AnsOX = sDoc_AnsOX	
	olect.frectsearchKey = searchKey
	olect.frectsearchString = searchString
	olect.fnGetlecturerList()
%>

<script type="text/javascript">

function goWrite(didx){
	location.href = "lecturer_read.asp?didx="+didx+"&<%=vParam%>&menupos=<%=g_MenuPos%>";
}

function goedit(didx){
	location.href = "lecturer_write.asp?didx="+didx+"&<%=vParam%>&menupos=<%=g_MenuPos%>";
}

function godel(didx){
	
	if (confirm('���� ���� �Ͻðڽ��ϱ�?')){
	var godel = window.open('lecturer_proc.asp?didx='+didx+'&mode=del','godel','width=600,height=400,scrollbars=yes,resizable=yes');
	godel.focus();
	}
}

function reg(){
	frm.Statusgubun.value='ON'
	frm.submit
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="get">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<input type="hidden" name="Statusgubun" value="<%=Statusgubun%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ó������:
		<%=CommonCode("w","K000",sDoc_Status)%>		
     	��û����:
		<%=CommonCode("w","G000",sDoc_Type)%>
     	�亯����:
     	<select name="ans_ox" class='select'>
	     	<option value='' selected>��ü</option>
	     	<option value='x' <% If sDoc_AnsOX = "x" Then %>selected<% End If %>>�̴亯</option>
	     	<option value='o' <% If sDoc_AnsOX = "o" Then %>selected<% End If %>>�亯�Ϸ�</option>
     	</select>
     	�˻�����:
     	<% DrawMainPosCodeCombo "searchKey" ,searchKey%>
     	<input type="text" name="searchString" size="20" value="<%= searchString %>">
     	<input type="submit" value="�˻�" class="button" onfocus="reg();">
     	<br>
     	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
	</td>
</tr>
</form>
</table>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�űԵ��" onClick="location.href='lecturer_write.asp?menupos=<%=g_MenuPos%>'">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if olect.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= olect.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= olect.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td>��ȣ</td>
	<td>�����</td>
	<td>����</td>		
	<td>�߿䵵</td>
	<td>�����</td>
	<td>��������</td>
	<td>ó������</td>
	<td>���ÿ���</td>
	<td>���</td>	
</tr>
<%
For i =0 To olect.fresultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');"><%=olect.FItemList(i).fdoc_idx%></td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%= getthefingers_staff("", olect.FItemList(i).fpart_sn, olect.FItemList(i).fcompany_name) %>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');" align="left"><%=olect.FItemList(i).fdoc_subject%></td>		
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');"><%=olect.FItemList(i).fdoc_important_nm%></td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<%=FormatDatetime(olect.FItemList(i).fdoc_regdate,2)%>
	</td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');"><%=olect.FItemList(i).fdoc_type_nm%></td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');"><%=olect.FItemList(i).fdoc_status_nm%></td>
	<td onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');">
		<% 
		if olect.FItemList(i).fans_count > 0 then 
			response.write olect.FItemList(i).fans_count & "��"
		else
			response.write olect.FItemList(i).fdoc_ans_ox
		end if
		%>	
	</td>
	<td>
		<%
		if olect.FItemList(i).fdoc_id = session("ssBctId") or (fingmaster) then
		%>
		<input type="button" onclick="goedit('<%=olect.FItemList(i).fdoc_idx%>');" value="����" class="button">
		<input type="button" onclick="godel('<%=olect.FItemList(i).fdoc_idx%>');" value="����" class="button">
		<% end if %>	
		<input type="button" onclick="goWrite('<%=olect.FItemList(i).fdoc_idx%>');" value="����" class="button">
	</td>	
</tr>
<%
Next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if olect.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= olect.StartScrollPage-1 %>&<%=vParam%>&menupos=<%=g_MenuPos%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + olect.StartScrollPage to olect.StartScrollPage + olect.FScrollCount - 1 %>
			<% if (i > olect.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(olect.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&<%=vParam%>&menupos=<%=g_MenuPos%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if olect.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&<%=vParam%>&menupos=<%=g_MenuPos%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<%
else
%>
<tr bgcolor="#FFFFFF" height="30">
	<td colspan="20" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
</tr>
<%
End If
%>
</table>

<%
set olect = nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->

<%
	''session.codePage = 949
%>