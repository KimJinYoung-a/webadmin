<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �Կ� ��û ���������
' History : 2012.03.13 ������ ����
'			2015.07.28 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
IF application("Svr_Info")="Dev" THEN
	g_MenuPos   = "1404"		'### �޴���ȣ ����.
Else
	g_MenuPos   = "1419"		'### �޴���ȣ ����.
End If

Dim lPhotoreq, page, i, makerid, cdl, r_use, s_type, num_name, req_status_type, request_name, req_photo_user, req_stylist
Dim iPageSize, iCurrentpage ,iDelCnt, sSearchTeam, sDoc_Status, sDoc_AnsOX, sSearchMine, confirmdate, tmpconfirmdate, j
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim iTotCnt
	page = request("page")

If page = "" Then page = 1

'�˻� Get�� ��..
makerid 		= request("makerid")
cdl				= request("req_category")
r_use			= request("req_use")
s_type			= request("s_type")
num_name 		= request("num_name")
req_status_type = request("req_status_type")
request_name 	= request("request_name")
req_photo_user 	= request("req_photo_user")
req_stylist		= request("req_stylist")

set lPhotoreq = new Photoreq
	lPhotoreq.FPageSize = 20
	lPhotoreq.FCurrPage = page
	lPhotoreq.FMakerid = makerid
	lPhotoreq.FCdl = cdl
	lPhotoreq.FReq_use = r_use
	lPhotoreq.FS_type = s_type
	lPhotoreq.FNum_Name = num_name
	lPhotoreq.FReq_status_type = req_status_type
	lPhotoreq.FRequest_name = request_name
	lPhotoreq.FReq_photo_user = Trim(req_photo_user)
	lPhotoreq.FReq_stylist = Trim(req_stylist)
	lPhotoreq.fnPhotoreqlist
%>
<script language="javascript">
function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=410,height=600');
}
function user_manage()
{
	window.open('PopUserList.asp','coopcode','width=410,height=600');
}
function gosubmit(page){
    document.searchfrm.page.value=page;
	document.searchfrm.submit();
}
function goUpdate(didx)
{
	location.href = "/admin/photo_req/request_modi.asp?req_no="+didx+"&udate=A&menupos=<%= menupos %>";
}
</script>
<p>
<!-- height="100%" �̰� �������� ������ ������ �Ʒ� �Կ���û����Ʈ�� �ȳ���. ������. -->
<iframe src="/admin/photo_req/board_list.asp" name="board" width="100%" height="200" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<p>
<!-- //-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr><td><b>[�Կ���û����Ʈ]</b></td></tr>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="searchfrm" action="request_list.asp" method="get">
	<tr align="center" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("gray") %>" width="100">�˻�����</td>
		<td align="left">
			<table width="100%" align="center"  cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="FFFFFF">
				<td width="150">�귣�� : </td>
				<td><%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
				<td width="150">����ī�װ� : </td>
				<td colspan="5">
					<%' call DrawCategoryLarge_disp("req_category", cdl) %>
					<%= fnStandardDispCateSelectBox(1,cdl, "req_category", cdl, "")%>
				</td>
			</tr>
			<tr bgcolor="FFFFFF">
				<td width="150">�Կ��뵵 : </td>
				<td><% call DrawPicGubun("req_use", r_use, "2") %></td>
				<td width="150">no/��ǰ�� : </td>
				<td colspan="5">
					<select name="s_type" class="select">
						<option value="">--no/��ǰ����--</option>
						<option value="1" <%If s_type = "1" Then response.write "selected" End If%>>��û�� no</option>
						<option value="2" <%If s_type = "2" Then response.write "selected" End If%>>��ǰ��</option>
					</select>
					<input type="text" class="text" name="num_name" value="<%=num_name%>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();">
				</td>
			</tr>
			<tr bgcolor="FFFFFF">
				<td width="150">������� : </td>
				<td>
					<select name="req_status_type" class="select">
						<option value="">--������¼���--</option>
						<option value="4" <%If req_status_type = "4" Then response.write "selected" End If%>>�߰����� ��û</option>
						<option value="1" <%If req_status_type = "1" Then response.write "selected" End If%>>�Կ������� ����</option>
						<option value="2" <%If req_status_type = "2" Then response.write "selected" End If%>>�Կ���</option>
						<option value="3" <%If req_status_type = "3" Then response.write "selected" End If%>>�Կ��Ϸ�</option>
						<option value="9" <%If req_status_type = "9" Then response.write "selected" End If%>>��������</option>
					</select>
				</td>
				<td width="150">�Կ���û�� : </td>
				<td><input type="text" class="text" name="request_name" size="16" maxlength="16" value="<%=request_name%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
				<td>������� : </td>
				<td><input type="text" class="text" name="req_photo_user" size="16" maxlength="16" value="<%=req_photo_user%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
				<td>��罺Ÿ�ϸ���Ʈ : </td>
				<td><input type="text" class="text" name="req_stylist" size="16" maxlength="16" value="<%=req_stylist%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
			</tr>
			</table>
		</td>
		<td bgcolor="<%= adminColor("gray") %>" width="100"><input type="button" class="button_s" value="�˻�" onClick="javascript:document.searchfrm.submit();"></td>
	</tr>
</table>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="���ε��" onClick="location.href='request_write.asp?menupos=<%=menupos%>&iC=<%=iCurrentpage%>'">
		<input type='button' class='button' value='����' onClick='user_manage()'>
		&nbsp;<font color="red"><ins>* ��� �� ���� ��������,�ʵ� ��Ź �帳�ϴ�.</ins></font>
	</td>
	<td align="right">
		<%
			Response.Write "<input type='button' class='button' value='�ڵ����' onClick='code_manage()'>&nbsp;"
		%>
	</td>
</tr>
</table>
<br>
<!-- �׼� �� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">�˻���� : <b><%= lPhotoreq.FTotalCount %></b>&nbsp;&nbsp;&nbsp;&nbsp;������ : <b><%=page%>/<%=lPhotoreq.FTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="60">��û��No</td>
	<td width="100">�������</td>
	<td width="200">�Կ��뵵</td>
	<td width="">��ǰ��(��ȹ����)</td>
	<td width="100">ī�װ�</td>
	<!--<td width="">�귣��</td>-->
	<td width="130">��û�Ͻ�</td>
	<td width="260">�Կ�Ȯ���Ͻ�</td>
	<td width="60">���MD<BR>(�Կ���û)</td>
	<td width="60">�߿䵵</td>
	<td width="50">�ϼ�URL<Br>��Ͽ���</td>
</tr>
<%
	If lPhotoreq.FResultcount = 0 Then
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="10" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
	</tr>
<%
	Else
		For i = 0 to lPhotoreq.FResultcount -1
%>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="goUpdate('<%=lPhotoreq.FPhotoreqList(i).FReq_no%>')">
		<td><%=lPhotoreq.FPhotoreqList(i).FReq_no%></td>
		<td>
			<%
			Select Case lPhotoreq.FPhotoreqList(i).FReq_status
				Case "0"	lPhotoreq.FPhotoreqList(i).FReq_status  = ""
				Case "1"	lPhotoreq.FPhotoreqList(i).FReq_status  = "�Կ������� ����"
				Case "2"	lPhotoreq.FPhotoreqList(i).FReq_status  = "�Կ���"
				Case "3"	lPhotoreq.FPhotoreqList(i).FReq_status  = "�Կ��Ϸ�"
				Case "4"	lPhotoreq.FPhotoreqList(i).FReq_status  = "�߰� ���� ��û��"
				Case "9"	lPhotoreq.FPhotoreqList(i).FReq_status  = "��������"
			End Select

			Select Case lPhotoreq.FPhotoreqList(i).FFontColor
				Case "R"	response.write "<font color='RED'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
				Case "G"	response.write "<font color='GREEN'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
				Case Else	response.write "<font color='BLACK'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
			End Select
			%>
		</td>
		<td>
			<%=lPhotoreq.FPhotoreqList(i).FReq_use%>
			<%
				If lPhotoreq.FPhotoreqList(i).FReq_use_detail <> "" Then
					response.write "("&lPhotoreq.FPhotoreqList(i).FReq_use_detail&")"
				End If
			%>
		</td>
		<td align="left"><%=DDotFormat(lPhotoreq.FPhotoreqList(i).FReq_prd_name,20)%></td>
		<td><%=lPhotoreq.FPhotoreqList(i).FReq_codenm%></td>
		<!--<td><%'=lPhotoreq.FPhotoreqList(i).FReq_makerid%></td>-->
		<td>
			��û �Ͻ� : <%= Left(lPhotoreq.FPhotoreqList(i).FReq_regdate,10) %><br>
		</td>
		<td>
			<%
			confirmdate = lPhotoreq.FPhotoreqList(i).fconfirmdate
			%>
			<% if confirmdate <> "" then %>
				<% For j = LBound(Split(confirmdate,"|^|")) To UBound(Split(confirmdate,"|^|")) %>
				<%
				tmpconfirmdate = Split(confirmdate,"|^|")(j)
				tmpconfirmdate = Split(tmpconfirmdate,"|*|")
				%>
				<%= left(tmpconfirmdate(0),10) %>
				<% if tmpconfirmdate(2)<>"" or tmpconfirmdate(3)<>"" then %>
					(
					<% if tmpconfirmdate(2) <> "" then %>
						���� : <%= tmpconfirmdate(2) %>
					<% end if %>
					<% if tmpconfirmdate(3) <> "" then %>
						, ��Ÿ�� : <%= tmpconfirmdate(3) %>
					<% end if %>
					)
				<% end if %>
				<br>
				<% next %>
			<% end if %>
		</td>
		<td>
			<%
			If isnull(lPhotoreq.FPhotoreqList(i).FMDid) = "False" Then
				response.write lPhotoreq.FPhotoreqList(i).FMDid&"<br>("& lPhotoreq.FPhotoreqList(i).FReq_name &")"
			ElseIf isnull(lPhotoreq.FPhotoreqList(i).FMDid) = "True" or (lPhotoreq.FPhotoreqList(i).FMDid) = "00" Then
				response.write lPhotoreq.FPhotoreqList(i).FReq_name
			End If
			%>
		</td>
		<td>
			<% for j = 1 to lPhotoreq.FPhotoreqList(i).FImport_level %>��<% next %>
		</td>
		<td>
			<% if lPhotoreq.FPhotoreqList(i).fopencount>0 then %>
				Y
			<% else %>
				N
			<% end if %>
		</td>
	</tr>
<%
		Next
	End If
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lPhotoreq.HasPreScroll Then %>
			<a href="javascript:gosubmit('<%= lPhotoreq.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lPhotoreq.StartScrollPage to lPhotoreq.StartScrollPage + lPhotoreq.FScrollCount - 1 %>
			<% If (i > lPhotoreq.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lPhotoreq.FCurrPage) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');">[<%= i %>]</a>
			<% End if %>
		<% Next %>
		<% If lPhotoreq.HasNextScroll Then %>
			<a href="javascript:gosubmit('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</form>
</table>
<%set lPhotoreq = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->