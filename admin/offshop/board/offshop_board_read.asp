<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ���� �Խ���
' History : 2010.06.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls.asp"-->

<%
Dim arrFileList, i , olect , lectFile , oread,dispshopall ,dispshopdiv ,shopidcount ,doc_kind
dim sDoc_ViewList ,sDoc_WorkerView ,dispshopidon ,oshop ,dispshopdivon
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Type, sDoc_Import, sDoc_Subj, sDoc_Content
Dim sDoc_UseYN, sDoc_Regdate , vParam, s_status, s_type, s_ans_ox , g_MenuPos ,sDoc_WorkerName
	iDoc_Idx		= requestCheckVar(Request("didx"),10)
	s_status		= requestCheckVar(Request("s_status"),10)
	s_type			= requestCheckVar(Request("s_type"),10)
	s_ans_ox		= requestCheckVar(Request("s_ans_ox"),1)
	g_MenuPos = requestCheckVar(request("menupos")	,10)		
		
	'�д� �������� ���뵵 ������ �Ǿ Ȥ�ó� �� �Ͽ� ����Ͽ� �Ķ���͸��� �ٲ㼭 �ְ� �޾ҽ�
	vParam = "doc_status="&s_status&"&doc_type="&s_type&"&ans_ox="&s_ans_ox
	
If iDoc_Idx = "" Then
	sDoc_Id 		= session("ssBctId")
	sDoc_Name		= session("ssBctCname")
	sDoc_Regdate	= Left(now(),10)
	sDoc_WorkerName	= ""
	sDoc_Worker		= ""
Else
	
	'//������
	Set olect = New clecturer_list
		olect.FrectDoc_Idx = iDoc_Idx
		olect.fnGetlecturerView()
	
		sDoc_Id 		= olect.FOneItem.FDoc_Id
		sDoc_Name		= olect.FOneItem.Fusername
		sDoc_Status		= olect.FOneItem.FDoc_Status
		if sDoc_Status = "" then sDoc_Status = "K001"	
		sDoc_Type		= olect.FOneItem.FDoc_Type
		sDoc_Import		= olect.FOneItem.FDoc_Import
		sDoc_Subj		= ReplaceBracket(olect.FOneItem.FDoc_Subj)
		sDoc_Content	= ReplaceBracket(olect.FOneItem.FDoc_Content)
		sDoc_UseYN		= olect.FOneItem.FDoc_UseYN
		sDoc_Regdate	= olect.FOneItem.FDoc_Regdate
		shopidcount = olect.FOneItem.fshopidcount
		dispshopall = olect.FOneItem.fdispshopall
		dispshopdiv = olect.FOneItem.fdispshopdiv
		doc_kind = olect.FOneItem.fdoc_kind
		
		if shopidcount > 0 then dispshopidon = "ON"
		if dispshopdiv <> "" and not isnull(dispshopdiv) then dispshopdivon = "ON"
		
	'/�� Ȯ�� ��¥ ����		'/�����̸� ����
	if session("ssBctId") <> sDoc_Id then
		Call WorkerView(iDoc_Idx)
	end if

	'//�� Ȯ���� ��¥ ����Ʈ
	Set oread = New clecturer_list
		oread.FrectDoc_Idx = iDoc_Idx
		oread.fnGetlecturerread()

		sDoc_WorkerName	= oread.FDoc_WorkerName
		sDoc_WorkerView	= oread.FDoc_WorkerViewdate	

	'/÷������ ����Ʈ
	set lectFile = new clecturer_list
	 	lectFile.FrectDoc_Idx = iDoc_Idx
		arrFileList = lectFile.fnGetFileList()

	'//��Ź���� ����Ʈ
    set oshop = new clecturer_list
    oshop.FrectDoc_Idx = iDoc_Idx
    
    '/Ư�������� ������쿡�� ����
    if shopidcount > 0 then
    	oshop.getShopList
    end if
    
	'/Ȯ������ �ִ°�쿡��
	For i=0 To UBOUND(Split(sDoc_WorkerName,","))
		if Not(sDoc_WorkerView="" or isNull(sDoc_WorkerView)) then
			sDoc_ViewList = sDoc_ViewList & "&nbsp;" & Split(sDoc_WorkerName,",")(i) & " : " & Split(sDoc_WorkerView,",")(i) & "<br>"
		end if
	Next		
End If
%>

<script type='text/javascript'>

function checkform(frm)
{
	frm.submit();
}

</script>

�غ������ : ���翡�� �� ���忡 �˸��� �������� �Դϴ�.(���翡���� �ۼ�����)
<br>�������� : �����忡�� ���翡 ��û�ϴ� ���Դϴ�.(���忡���� �ۼ�����)
<br>������� : �����忡�� Ÿ ���忡 �˸��� ���������Դϴ�.(���忡���� �ۼ�����)
<form name="frm" action="offshop_board_proc.asp" method="post">
<input type="hidden" name="didx" value="<%=iDoc_Idx%>">
<input type="hidden" name="mode" value="view">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10"> 
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<% If iDoc_Idx <> "" Then %>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��ȣ</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=iDoc_Idx%></td>
			</tr>
			<% End If %>
			<input type="hidden" name="doc_useyn" value="Y">
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�����</td>
				<td width="700" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Name%>(<%=sDoc_Id%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �����: <%=sDoc_Regdate%></td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=CommonCode("v","G000",sDoc_Type,C_ADMIN_USER,"")%>
					<% if sDoc_Type = "02" then %>
						(���� ���� : <%=CommonCode("w","K000",sDoc_Status,"","")%>)
					<% end if %>
					<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
				</td>
			</tr>
			<tr >
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��������</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<% if sDoc_Type <> "02" then %>
						<% if dispshopall="Y" then %>
							��ü����<Br>
						<% end if %>
						<% if dispshopdivon="ON" then %>
							<%=CommonCode("v","A000",dispshopdiv,"","")%><Br>
						<% end if %>
						<% if dispshopidon="ON" then %>
							Ư������<Br>
					        <%
					        if iDoc_Idx <> "" then
					        	if oshop.FResultCount > 0 then
					        
					        	for i=0 to oshop.FResultCount-1
					        %>
								&nbsp;&nbsp;&nbsp;&nbsp;-<%= oshop.FItemList(i).fshopname %><Br>
					        <%
					        	next
					        
					    		end if
					    	end if
					        %>
						<% end if %>
					<% else %>
						����
					<% end if %>
				</td>
			</tr>
			<tr >
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=CommonCode("v","doc_kind",doc_kind,"","") %>
				</td>
			</tr>			
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�߿䵵</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=CommonCode("v","L000",sDoc_Import,"","")%>						
				</td>
			</tr>
			<input type="hidden" name="doc_difficult" value="2">

			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDoc_Subj%>
				</td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�� ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<%=sDoc_Content%>
				</td>
			</tr>
			<tr height="30">
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">÷������</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
					<table cellpadding="0" cellspacing="0" border="0" class="a">
					<tr>
						<td width="100%" style="padding:3 0 3 10">
							<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
							<%
							IF isArray(arrFileList) THEN
								For i =0 To UBound(arrFileList,2)
							%>
								<tr>
									<td>											
										<a href='<%=arrFileList(0,i)%>' target='_blank'>
										<%=Split(Replace(arrFileList(0,i),"http://",""),"/")(4)%></a>
									</td>
								</tr>
							<%
								Next
								Response.Write "<input type='hidden' name='isfile' value='o'>"
							Else
							%>
								<tr>
									<td>
									</td>
								</tr>
							<% End If %>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<% if sDoc_ViewList <> "" then %>
				<tr >
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��Ȯ��<br>���</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<%= sDoc_ViewList %>
					</td>
				</tr>
			<% end if %>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<% If iDoc_Idx <> "" AND sDoc_Id = session("ssBctId") and sDoc_Type = "02" Then %>	
			<input type="button" onclick="checkform(frm);" value="�����ϱ�" class="button">
		<% end if %>
		<input type="button" value="�������" onclick="location.href='offshop_board.asp?menupos=<%=g_MenuPos%>'" class="button">		
	</td>	
</tr>
</table>
</form>

<% If iDoc_Idx <> "" Then %>
<!-- ####### �亯���� ####### //-->
<br>
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr height="30">
	<td>
		<img src="/images/icon_arrow_link.gif"></td><td style="padding-top:3">&nbsp;<b>�亯</b>
	</td>
</tr>
</table>
<iframe src="iframe_board_ans.asp?didx=<%=iDoc_Idx%>&registid=<%=sDoc_Id%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### �亯���� ####### //-->
<% End If %>

<%	
	set olect = nothing
	set lectFile = nothing
	set oread = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->