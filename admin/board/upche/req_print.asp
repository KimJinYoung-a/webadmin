<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<%
'###########################################################
' Description : ��ü ��������
' History : 2008.09.01 �ѿ�� ����/�߰�
'###########################################################
%>
<%
dim i, j
dim commmode
	commmode=request("commmode")
dim page,gubun, onlymifinish
dim research, searchkey,catevalue
dim ipjumYN
	page = request("pg")
	gubun = request("gubun")
	onlymifinish = request("onlymifinish")
	research = request("research")
	searchkey = request("searchkey")
	catevalue=request("catevalue")
	ipjumYN=request("ipjumYN")
	if research="" and onlymifinish="" then onlymifinish="on"

	'// �⺻������ �����Ƿڼ�
	if gubun="" then gubun="01"
	if (page = "") then page = "1"

dim companyrequest
	set companyrequest = New CCompanyRequest
	companyrequest.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
	A:link, A:visited, A:active { text-decoration: none; }
	A:hover { text-decoration:underline; }
	BODY, TD, UL, OL, PRE { font-size: 10pt; }
	INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #ffffff; color: #000000; }
-->
</STYLE>

<!-- ��ü���� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="black">
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<tr bgcolor="FFFFFF" align="center">
		<td colspan=5><b><font size=3 color="blue">���¾�ü ���� ��������</font></b></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">ȸ���</td>	
		<td><%= db2html(companyrequest.results(0).companyname) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">��ǥ�ڸ�</td>			
		<td><%= db2html(companyrequest.results(0).chargename) %></td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">�ּ�</td>	
		<td><%= db2html(companyrequest.results(0).address) %></td>
		<td bgcolor="<%= adminColor("gray") %>">���Ű�</td>	
		<td>
			<%= db2html(companyrequest.results(0).cur_target) %>
		</td>							
	</tr>		
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">����</td>			
		<td><%= db2html(companyrequest.results(0).chargename) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">��å(�μ���)</td>			
		<td><%= db2html(companyrequest.results(0).chargeposition) %></td>			
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">Tel</td>			
		<td><%= db2html(companyrequest.results(0).phone) %></td>	
		<td bgcolor="<%= adminColor("gray") %>">H.P</td>			
		<td><%= db2html(companyrequest.results(0).hp) %></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">�����<br>��Ϲ�ȣ</td>	
		<td><%= db2html(companyrequest.results(0).license_no) %></td>
		<td bgcolor="<%= adminColor("gray") %>">�̸���</td>			
		<td><a href="mailto:<%= db2html(companyrequest.results(0).email) %>"><%= db2html(companyrequest.results(0).email) %></a></td>	
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">����</td>	
		<td>
			<% 
			if companyrequest.results(0).Service <> "" then
				if left(companyrequest.results(0).Service,1) <> 0 then response.write "����. "
				if mid(companyrequest.results(0).Service,3,1) <> 0 then response.write "����. "
				if mid(companyrequest.results(0).Service,5,1) <> 0 then response.write "�Ҹ�. "	 
				if mid(companyrequest.results(0).Service,7,1) <> 0 then response.write "����. "
				if mid(companyrequest.results(0).Service,9,1) <> 0 then response.write "����. "
				if mid(companyrequest.results(0).Service,11,1) <> 0 then response.write "����. "	
				if right(companyrequest.results(0).Service,1) <> 0 then response.write "��Ÿ. "
			end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">��ǰ��</td>	
		<td>
			<% Drawcatelarge "catelargebox",companyrequest.results(0).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">����</td>	
		<td>
			<% 
			if companyrequest.results(0).physical = 0 then 
				response.write "�����ü� ��ü����"
				response.write "("& companyrequest.results(0).physical_name & ")"
			else 
				response.write "����������ü Ư��"
				response.write "("& companyrequest.results(0).physical_name & ")"
			end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">����</td>	
		<td>
			<% 
			if companyrequest.results(0).manufacturing = 0 then 
				response.write "������� ��ü����"
				response.write "("& companyrequest.results(0).manufacturing_name & ")"
			else 
				response.write "�ܺξ�ü Ư��"
				response.write "("& companyrequest.results(0).manufacturing_name & ")"
			end if
			%>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">������� ���</td>	
		<td>
			<%= companyrequest.results(0).industrial %>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">���̼��� ���</td>	
		<td>
			<%= companyrequest.results(0).license %>
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">������</td>	
		<td>
			<% 
			if left(companyrequest.results(0).utong,1) <> 0 then response.write "�������Ǹ� "
			if mid(companyrequest.results(0).utong,3,1) <> 0 then response.write "��ȭ�� "
			if mid(companyrequest.results(0).utong,5,1) <> 0 then response.write "������ "	 
			if mid(companyrequest.results(0).utong,7,1) <> 0 then response.write "�븮�� "
			if mid(companyrequest.results(0).utong,9,1) <> 0 then response.write "���ռ��θ� "
			if mid(companyrequest.results(0).utong,11,1) <> 0 then response.write "Ȩ���� "	
			if mid(companyrequest.results(0).utong,13,1) <> 0 then response.write "Ÿ��ŷ�ó "
			if mid(companyrequest.results(0).utong,15,1) <> 0 then response.write "�ڻ�� "	
			if right(companyrequest.results(0).utong,1) <> 0 then response.write "�ڻ�� "				
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">���������</td>	
		<td>
			<% 
			if companyrequest.results(0).tax = 0 then 
				response.write "���� "
			elseif  companyrequest.results(0).tax = 1 then 
				response.write "�鼼 "
			elseif  companyrequest.results(0).tax = 2 then 
				response.write "�Ϲ� "			
			else
				response.write "���� "
			end if
			%>
		</td>				
	</tr>	
	<tr bgcolor="FFFFFF" align="center">
		<td bgcolor="<%= adminColor("gray") %>">ȸ��URL</td>	
		<td>
			<%
				dim arrUrl
				arrUrl = split(companyrequest.results(0).companyurl,",")
				if ubound(arrUrl)>0 then
					Response.Write "<a href='"
					if Left(arrUrl(0),7) <> "http://" then Response.Write "http://"
					Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
					Response.Write "<br><br><b>�������θ�</b> : " & arrUrl(1)
				else
					Response.Write "<a href='"
					if Left(companyrequest.results(0).companyurl,7) <> "http://" then Response.Write "http://"
					Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
				end if
			%>
		</td>			
		<td bgcolor="<%= adminColor("gray") %>">����</td>	
		<td>
			<%= companyrequest.code2name(companyrequest.results(0).reqcd) %>
		</td>				
	</tr>		
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">��ǰ��(�귣���)</td>	
		<td colspan=3>
			<%= db2html(companyrequest.results(0).reqcomment) %>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">÷������</td>	
		<td>
			<% if (companyrequest.results(0).attachfile <> "") then %>
				<a href="http://imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">�ٿ�ޱ�</a>
			<% else %>
				����
			<% end if %>
		</td>						
		<td bgcolor="<%= adminColor("gray") %>">ó������</td>	
		<td>
			<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
				�̿Ϸ�
			<% else %>
				<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
			<% end if %>
		</td>
	</tr>		
	<tr bgcolor="FFFFFF" align="center">			
		<td bgcolor="<%= adminColor("gray") %>">ȸ�缳��</td>	
		<td colspan=3>
			<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
		</td>				
	</tr>
	<tr bgcolor="FFFFFF" align="center">			
		<td colspan=4 align="left">
		<input type="button" value="����Ʈ" class="button" onclick="javascript:window.print();">
		</td>				
	</tr>	
</table><br>
<!-- ��ü���� �� -->

<!-- #include virtual="/lib/db/dbclose.asp" -->