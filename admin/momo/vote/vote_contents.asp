<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� �츮������������
' Hieditor : 2009.11.20 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i
dim idx,vote_num,contents_num,contents,isusing,regdate
	vote_num = requestcheckvar(request("vote_num"),8)

'//��
set ocontents = new cvote_list
	ocontents.FPageSize = 50
	ocontents.FCurrPage = 1
	ocontents.frectvote_num = vote_num
	
	'//���� ������쿡�� ����
	if vote_num <> "" then
	ocontents.fvote_contents()
	end if
%>

<script language="javascript">
	
	//��ǥ�߰�
	function insertvote(){		
		document.all.div1.innerHTML	= document.all.div1.innerHTML + "<input type='text' name='contents' size=64 maxlength=64><br>";		
	}

	//����
	function reg(){
		var nm  = document.getElementsByName('contents');
		
		if (nm.length==0){
			alert('��ǥ������ �Է��ϼ���');
			return;
		}
			
		for(var i=0 ; i < nm.length ; i++ ){		
			if (nm[i].value==''){
				alert('��ǥ������ �Է��ϼ���');
				nm[i].focus();
				return;
			}
		
		}
		
		frm.action='/admin/momo/vote/vote_process.asp';
		frm.mode.value='contents';
		frm.submit();
	}
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>voteid</td>
	<td bgcolor="#FFFFFF" align="left">
		<%= vote_num %><input type="hidden" name="vote_num" value="<%= vote_num %>">		
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǥ�ۼ�</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="button" onclick="insertvote();" value="��ǥ�߰�" class="button">
		<% 
		if ocontents.ftotalcount > 0 then
			for i = 0 to ocontents.ftotalcount - 1
		%>
			<input type='text' name='contents' value="<%=ocontents.fitemlist(i).fcontents%>" size=64 maxlength=64><br>
		<%
			next
		end if
		%>	
		<div name="div1" id="div1">
				
		</div>
	</td>
</tr>
<tr align="center" bgcolor="FFFFFF">
	<td colspan=2><input type="button" onclick="reg();" value="����" class="button"></td>
</tr>
</form>
</table>

<%
	set ocontents = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
