<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
'####################################################
' Description : �Կ� ��û ���� & �� ������
' History : 2012.03.15 ������ ����
'####################################################

	Dim gub, gubnm, i, udate
	Dim cPhotoreq, rno, arrFileList, sMode2
	Dim PhotoCnt
	
	rno = request("req_no")
	gub = request("gub")

	set cPhotoreq = new Photoreq
		cPhotoreq.FReq_no = rno
		cPhotoreq.fnPhotoreqUpdate
	If cPhotoreq.FPhotoreqList(0).FReq_use = "" Then
		Call Alert_move("�ش� ������ �����ϴ�","request_list.asp")
	End If
%>
<script language="Javascript">
<!--
function printpage() {
	window.print();
}
//-->
</script>
<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="25">
    <td align="left"><font size="3"><strong>�Կ���û��_<%=cPhotoreq.FPhotoreqList(0).FReq_department%> / 
<%
		If isnull(cPhotoreq.FPhotoreqList(0).FMDid) = "False" Then
			response.write cPhotoreq.FPhotoreqList(0).FMDid&"("& cPhotoreq.FPhotoreqList(0).FReq_name &")"
		ElseIf isnull(cPhotoreq.FPhotoreqList(0).FMDid) = "True" or (cPhotoreq.FPhotoreqList(0).FMDid) = "00" Then
			response.write cPhotoreq.FPhotoreqList(0).FReq_name
		End If
%>
	   	/ <%=cPhotoreq.FPhotoreqList(0).FPrd_name%></strong>
    </td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="25">
    <td align="right">�߿䵵 : <%=cPhotoreq.FPhotoreqList(0).FImport_level%></td>
</tr>
</table>
<p>
<table width="100%" align="center"  border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<col width="15%"></col>
<col width="30%"></col>
<col width="17%"></col>
<col width="38%"></col>
<tr>
	<td height="30" bgcolor="#DDDDFF">�Կ���û�Ͻ�</td>
	<td><%=Left(cPhotoreq.FPhotoreqList(0).FReq_regdate,10)%></td>
	<td bgcolor="#DDDDFF">�Կ���û��</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_name%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">�Կ�Ȯ���Ͻ�</td>
	<td colspan="3">
<%
	If IsNull(cPhotoreq.FPhotoreqList(0).FStart_date) Then
		response.write "&nbsp;"
	Else	
		For i = 0 to cPhotoreq.FResultcount -1
%>
				<font color="BLUE">���� : <%=cPhotoreq.FPhotoreqList(i).FStart_date%></font> ~ <font color="RED">���� : <%=cPhotoreq.FPhotoreqList(i).FEnd_date%></font><br>
<%
		Next
	End If
%>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">����׷���</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_photoname%>&nbsp;</td>
	<td bgcolor="#DDDDFF">��Ÿ�ϸ���Ʈ</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FStylistname%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">�Կ�����</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_gubun%></td>
	<td bgcolor="#DDDDFF">�Կ��뵵</td>
	<td>
		<%=cPhotoreq.FPhotoreqList(0).FReq_use%>
		<%
			If cPhotoreq.FPhotoreqList(0).FReq_use_detail <> "" Then
				response.write "("&cPhotoreq.FPhotoreqList(0).FReq_use_detail&")"
			End If
		%>
	</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">��ǰ��(��ȹ����)</td>
	<td colspan="3"><%=cPhotoreq.FPhotoreqList(0).FPrd_name%></td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">��ǰ��/��</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FPrd_type%></td>
	<td bgcolor="#DDDDFF">�ǸŰ�</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FPrd_price&"��"%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">�귣��ID</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FMakerid%>&nbsp;</td>
	<td bgcolor="#DDDDFF">��û�μ�/ī�װ�</td>
	<td><%=cPhotoreq.FPhotoreqList(0).FReq_department%>/<%=cPhotoreq.FPhotoreqList(0).FReq_codenm%>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">�ʿ� �Կ���</td>
	<td colspan="3"><% call CheckBoxUseType("doc_use_type", rno, "3") %>&nbsp;</td>
</tr>
<tr>
	<td height="30" bgcolor="#DDDDFF">���� �Կ� ����</td>
	<td colspan="3"><% call CheckBoxUseType("doc_use_concept", rno, "4") %>&nbsp;</td>
</tr>
</table>
<Br>
<table width="100%" align="center"  border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td height="30" bgcolor="#DDDDFF">��ǰ Ư¡ �� �ֿ� ���� ����</td>
</tr>
<tr>
	<td height="30" valign="top"><%=replace(cPhotoreq.FPhotoreqList(0).FReq_etc1,vbCrLf,"<br>")%></td>
</tr>
</table><br>
<p>
���� ��ũ �� ���� url : <a href="<%=cPhotoreq.FPhotoreqList(0).FReq_url%>" target="_blank"><%=cPhotoreq.FPhotoreqList(0).FReq_url%></a>
<p><br>
<table width="100%" align="center" border="1" bordercolordark="White" bordercolorlight="black" class="a" cellpadding="0" cellspacing="0">
<tr>
	<td height="30" bgcolor="#DDDDFF">�Կ��� ���ǻ���</td>
</tr>
<tr>
	<td height="30" valign="top"><%=replace(cPhotoreq.FPhotoreqList(0).FReq_etc2,vbCrLf,"<br>")%>&nbsp;</td>
</tr>
</table>
<p>
<table width="100%">
<tr>
	<td align="right">
		<input type="button" class="button_s" value="�μ��ϱ�" onClick="printpage();">
	</td>
</tr>
</table>
</body>
<%set cPhotoreq = nothing%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->