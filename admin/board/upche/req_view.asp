<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��������
' History : ������ ����
'			2008.09.01 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/companyrequestcls.asp" -->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim i, j, arrUrl, commmode, page,gubun, onlymifinish, research, searchkey,catevalue,dispCate,maxDepth, ipjumYN, comreqID
dim companyrequest
	commmode=requestCheckvar(request("commmode"),10)
	page	 	= requestCheckvar(request("pg"),10)
	gubun 		= requestCheckvar(request("gubun"),2)
	onlymifinish= requestCheckvar(request("onlymifinish"),3)
	research 	= requestCheckvar(request("research"),3)
	searchkey 	= requestCheckvar(request("searchkey"),32)
	catevalue	= requestCheckvar(request("catevalue"),3)
	ipjumYN		= requestCheckvar(request("ipjumYN"),1)
	comreqID 	= requestCheckvar(request("id"),10)
	dispCate		= requestCheckVar(Request("disp"),16) 
	maxDepth		= 2

if research="" and onlymifinish="" then onlymifinish="on"

'// �⺻������ �����Ƿڼ�
if gubun="" then gubun="01"
if (page = "") then page = "1"


set companyrequest = New CCompanyRequest
	companyrequest.read(comreqID)

%> 
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function delreq(){
	if (confirm("�����Ͻðڽ��ϱ�?") ==true)
		frm.mode.value="reqdel"; 
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}

function SubmitForm(){
	if (confirm("ó�����¸� �Ϸ�� ��ȯ�մϱ�?") == true) { document.f.submit(); }
}
function catesubmit(){

	if (confirm("ī�װ��� ���� �մϴ�.") ==true)
		frm.mode.value="chcate"; 
		frm.disp.value=f.disp.value; 
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}
function sellsubmit(){

	if (confirm("�Ǹ������� �����մϴ�.") ==true)
		frm.mode.value="chsell";
		frm.sellgubun.value=f.sellgubun.value;
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}
function ipjumYNsubmit(){

	if(confirm("�������� �����մϴ�.") ==true)
		frm.mode.value="ipjum";
		frm.ipjumYN.value=f.ipjumYN.value;
		frm.action="/admin/board/upche/req_act.asp";
		frm.submit();
}

function sendmail(){
    var ireqmail = "<%= replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","") %>";

    if (ireqmail.length<2){
        alert('�����ּҰ� �ùٸ��� �ʽ��ϴ�.');
        return;
    }
    
	if(confirm("������ �����ðڽ��ϱ�?.") ==true)
	frmmail.submit();
}

function MovePage(page){
	frm.pg.value=page;
	frm.research.value="<%=research %>";
	frm.gubun.value="<%=gubun%>";
	frm.onlymifinish.value="<%=onlymifinish%>";
	frm.catevalue.value="<%=catevalue%>";
	frm.ipjumYNvalue="<%=ipjumYN%>";
	frm.searchkey.value="<%=searchkey%>";
	frm.action="/admin/board/upche/req_list.asp";
	frm.submit();
}
function editcomm(){
	frm.commmode.value="edit";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.action="/admin/board/upche/req_view.asp";
	frm.submit();
}
function savecomm(){
	frm.mode.value="comm";
	frm.id.value="<%= companyrequest.results(0).id %>";
	frm.user.value="<%= session("ssBctCname") %>";
	frm.comment.value=commfrm.comment.value;
	frm.action="/admin/board/upche/req_act.asp";
	frm.submit();
	}

function AddNewBrand(){
	var cate1 = document.f.disp.value;
	var popwin = window.open("/admin/member/addnewbrand_step1.asp?pcuserdiv=9999_02&companyno=<%= db2html(companyrequest.results(0).license_no) %>&hp=<%= db2html(companyrequest.results(0).hp) %>&email=<%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %>&cd1=<%= left(companyrequest.results(0).dispcate,3) %>&cate1="+cate1,"addnewbrand2","width=800 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�ذ���Ÿ - ��ü���Խ���
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��ü���� ���� -->
<form method="post" name="f" action="/admin/board/upche/req_act.asp" onsubmit="return false" style="margin:0px;">
<input type="hidden" name="mode" value="finish">
<input type="hidden" name="menupos" value="menupos">
<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
<table width="100%" align="center" cellpadding="5" cellspacing="1" bgcolor="black" class="a">
<tr bgcolor="FFFFFF">
	<td colspan=5><b><font color="blue">���¾�ü ���� ��������</font></b></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">ȸ���</td>	
	<td><%= db2html(companyrequest.results(0).companyname) %></td>	
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ǥ�ڸ�</td>			
	<td><%= db2html(companyrequest.results(0).chargename) %></td>				
</tr>

<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">�ּ�</td>	
	<td><%= db2html(companyrequest.results(0).address) %></td>
	<td bgcolor="<%= adminColor("gray") %>" align="center">���Ű�</td>	
	<td>
		<%= db2html(companyrequest.results(0).cur_target) %>
	</td>							
</tr>
	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">����</td>			
	<td><%= db2html(companyrequest.results(0).chargename) %></td>	
	<td bgcolor="<%= adminColor("gray") %>" align="center">��å(�μ���)</td>			
	<td><%= db2html(companyrequest.results(0).chargeposition) %></td>			
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">Tel</td>			
	<td><%= db2html(companyrequest.results(0).phone) %></td>	
	<td bgcolor="<%= adminColor("gray") %>" align="center">H.P</td>			
	<td><%= db2html(companyrequest.results(0).hp) %></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">����ڵ�Ϲ�ȣ</td>	
	<td><%= db2html(companyrequest.results(0).license_no) %></td>
	<td bgcolor="<%= adminColor("gray") %>" align="center">�̸���</td>			
	<td><a href="mailto:<%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %>"><%= db2html(replace(replace(replace(trim(companyrequest.results(0).email),"<br>",""),vbcrlf,""),"\n","")) %></a></td>	
</tr>
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">����</td>	
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
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">��ǰ��</td>	
	<td>
		<% Drawcatelarge "catelargebox",companyrequest.results(i).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)
	</td>				
</tr>	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">����</td>	
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
	<td bgcolor="<%= adminColor("gray") %>" align="center">����</td>	
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
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">������� ���</td>	
	<td>
		<%= companyrequest.results(0).industrial %>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">���̼��� ���</td>	
	<td>
		<%= companyrequest.results(0).license %>
	</td>				
</tr>	
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">������</td>	
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
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">���������</td>	
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
<tr bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("gray") %>" align="center">ȸ��URL</td>	
	<td>
		<%
			arrUrl = split(companyrequest.results(0).companyurl,",")
			if ubound(arrUrl)>0 then
				Response.Write "<a href='"
				if Left(arrUrl(0),7) <> "http://" and Left(arrUrl(0),8) <> "https://" then Response.Write "http://"
				Response.Write arrUrl(0) & "' target='_blank'>" & arrUrl(0) & "</a>"
				Response.Write "<br><br><b>�������θ�</b> : " & arrUrl(1)
			else
				Response.Write "<a href='"
				if Left(companyrequest.results(0).companyurl,7) <> "http://" and Left(companyrequest.results(0).companyurl,8) <> "https://" then Response.Write "http://"
				Response.Write companyrequest.results(0).companyurl & "' target='_blank'>" & companyrequest.results(0).companyurl & "</a>"
			end if
		%>
	</td>
		
	<td bgcolor="<%= adminColor("gray") %>" align="center">����</td>	
	<td>
		<%= companyrequest.code2name(companyrequest.results(0).reqcd) %>
	</td>				
</tr>
<tr bgcolor="FFFFFF">		
	<td bgcolor="<%= adminColor("gray") %>" align="center">��ǰ��(�귣���)</td>	
	<td colspan=3>
		<%= nl2br(db2html(companyrequest.results(0).reqcomment)) %>
	</td>
			
</tr>
<tr bgcolor="FFFFFF">		
	<td bgcolor="<%= adminColor("gray") %>" align="center">÷������</td>	
	<td>
		<% if (companyrequest.results(0).attachfile <> "") then %>
			<a href="//imgstatic.10x10.co.kr<%= companyrequest.results(0).attachfile %>" target="_blank">�ٿ�ޱ�</a>
		<% else %>
			����
		<% end if %>
	</td>
					
	<td bgcolor="<%= adminColor("gray") %>" align="center">ó������</td>	
	<td>
		<% if (IsNull(companyrequest.results(0).finishdate) = true) then %>
			�̿Ϸ�
		<% else %>
			<%= FormatDate(companyrequest.results(0).finishdate, "0000-00-00") %>
		<% end if %>
	</td>
</tr>
<tr bgcolor="FFFFFF"> 		
	<td bgcolor="<%= adminColor("gray") %>" align="center">ȸ�缳��</td>	
	<td colspan=3>
		<%= nl2br(db2html(companyrequest.results(0).companycomments)) %>
	</td>				
</tr>
<tr bgcolor="FFFFFF"> 			
	<td colspan=4 align="left">
	<input type="button" value="����Ʈ" class="button" onclick="javascript:window.print();">
	</td>				
</tr>
</table>
<!-- ��ü���� �� -->

<br>

<!-- ���� ����  ����-->
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">		
	<td colspan=3><b><font color="blue">��ü���� ����</font></b></td>							
</tr>
<tr bgcolor="FFFFFF">		
	<td width="100">ī�װ� ����</td>	
	<td width="300"><%if not isNull(companyrequest.results(0).dispcate) then%>
		<span><%=companyrequest.results(0).dispcatename1%> > <%=companyrequest.results(0).dispcatename2%></span>
		<%end if%>
		<div style="color:gray"> <%if companyrequest.results(0).cd1<>"" then%>����: <% Drawcatelarge "catelargebox",companyrequest.results(i).cd1 %>(<% Drawcatemid companyrequest.results(0).cd1,"catemidbox",companyrequest.results(0).cd2 %>)<%end if%></div>
	</td>					
	<td>
		<!-- #include virtual="/common/module/dispCateSelectBoxDepth.asp"-->  
		<input type=button value="����" onclick="catesubmit();">			
	</td>	
</tr>
<tr bgcolor="FFFFFF">		
	<td>�Ǹ����� ����</td>	
	<td>
		<% if companyrequest.results(0).sellgubun="Y" then %>
		ON-Line/OFF-Line
		<% elseif companyrequest.results(0).sellgubun="N" then%>
		ON-Line
		<% elseif companyrequest.results(0).sellgubun="F" then%>
		OFF-Line
		<% else %>
		��Ÿ
		<% end if %>
	</td>					
	<td>
		<select name="sellgubun" class="a">
			<option value="Y">ON-Line/OFF-Line</option>
			<option value="N">ON-Line</option>
			<option value="F">OFF-Line</option>
		</select>
		<input type=button value="����" onclick="sellsubmit();">			
	</td>	
</tr>
<tr bgcolor="FFFFFF">		
	<td>��������</td>	
	<td>
		<% if companyrequest.results(i).ipjumYN="Y" then response.write "�����Ϸ�" %>
		<% if companyrequest.results(i).ipjumYN="N" then response.write "������" %>
	</td>				
	<td>
	<!--<select name="ipjumYN" class="a">
		<option value="Y">���� �Ϸ�</option>
		<option value="N">�� ����</option>
	</select>-->
	<!--<input type=button value="����" onclick="ipjumYNsubmit();">-->
	</td>	
</tr>	
<tr bgcolor="FFFFFF">		
	<td colspan=3>
		<input type="button" value=" �Ϸ�ó�� " onclick="SubmitForm()" class="button">&nbsp;&nbsp;
		<% if companyrequest.results(i).fisusing="Y" then %>
			<input type="button" value="����" onclick="delreq()" class="button">&nbsp;&nbsp;
		<% end if %>
		<input type="button" value=" �������μ��� ������ " onclick="AddNewBrand()" class="button"> 
	</td>
</tr>
</table>
</form>
<!-- ���� ����  ��-->

<!-- �ڸ�Ʈ �κ� -->
<form name="commfrm" method="post" action="" onsubmit="return false" style="margin:0px;">
<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan=3><b><font color="blue">��ü���� ���Ϻ�����</font></b></td>
</tr>

<% if commmode="" and companyrequest.results(0).replyuser <>"" then %>
	<tr bgcolor="FFFFFF">
		<td width="10%" valign="top">
		�ۼ�: <%= db2html(companyrequest.results(0).replyuser) %>
		</td>
		<td width="75%" valign="top">
		<%= nl2br(db2html(companyrequest.results(0).replycomment)) %>
		</td>
		<td width="15%">
		<input type="button" value="����" onclick="javascript:editcomm();">
		</td>
	</tr>
	<tr bgcolor="FFFFFF" align="left">
		<td colspan=3><input type="button" value="mail������" onclick="javascript:sendmail();">	</td>
	</tr>

<% 
'//�������
elseif commmode="edit" then
%>
	<tr bgcolor="FFFFFF">
		<td width="10%" valign="top">
			�ۼ�: <%= session("ssBctCname") %>
		</td>
		<td valign="top">
			<textarea name="comment" rows=10 cols=95><%= db2html(companyrequest.results(0).replycomment) %></textarea>
		</td>
		<td>
			<input type="button" value="����" onclick="javascript:savecomm();">
		</td>
	</tr>
	
<% 
'//�ۼ����
elseif companyrequest.results(0).replyuser ="" then
%>
	<tr bgcolor="FFFFFF">
		<td valign="top">
			�ۼ�: <%= session("ssBctCname") %>
		</td>
		<td valign="top">
			<textarea name="comment" rows=10 cols=95></textarea>
		</td>
		<td>
			<input type="button" value="����" onclick="javascript:savecomm();">
		</td>
	</tr>
<% end if %>

</table>
</form>

<form name="frm" method="post" action="" onsubmit="return false" style="margin:0px;">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
	<input type="hidden" name="pg" value="<%= page %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="disp" value=""> 
	<input type="hidden" name="sellgubun" value="">
	<input type="hidden" name="ipjumYN" value="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="gubun" value="<%= gubun%>" >
	<input type="hidden" name="onlymifinish" value="<%=onlymifinish%>">
	<input type="hidden" name="catevalue" value="<%=catevalue%>">
	<input type="hidden" name="searchkey" value="<%=searchkey%>">
	<input type="hidden" name="commmode" value="">
	<input type="hidden" name="user" value="">
	<input type="hidden" name="comment" value="">
</form>
<form name="frmmail" method="post" action="/admin/board/upche/req_mail.asp" onsubmit="return false" style="margin:0px;">
	<input type="hidden" name="user" value="<%= session("ssBctCname") %>">
	<input type="hidden" name="userid" value="<%= session("ssBctId") %>">
	<input type="hidden" name="mailname" value="<%= companyrequest.results(0).chargename %>">
	<input type="hidden" name="mailto" value="<%= companyrequest.results(0).email %>">
	<input type="hidden" name="content" value="<%= companyrequest.results(0).replycomment %>">
	<input type="hidden" name="id" value="<%= companyrequest.results(0).id %>">
</form>
 
<script type="text/javascript">

//��ī�װ����ý� ��ī�װ� ����
function searchCD2(paramCodeLarge) {
		
	resetLeftCountrySelect() ;		
	resetLeftCitySelect() ;
	
	if(paramCodeLarge != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeLarge + "&form_name=f&element_name=cd2";
	}
}

//��ī�װ� ���ý� ��ī�װ� ����	
function searchCD3(paramCodeMid) {	
	resetLeftCitySelect() ;
	
	if(paramCodeMid != '') {
		FrameSearchCategory.location.href="/admin/CategoryMaster/frame_category_select.asp?search_code=" + paramCodeMid + "&form_name=f&element_name=cd3";
	}	 
}

//��ī�װ� �ʱ�ȭ
function resetLeftCountrySelect() {
	document.f.cd2.length = 1;
	document.f.cd2.selectedIndex = 0 ;
}

		
//��ī�װ� �ʱ�ȭ
function resetLeftCitySelect() {
	document.f.cd3.length = 1;
	document.f.cd3.selectedIndex = 0 ;
}

</script>

<%
set companyrequest=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->