<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ ��÷���
' History : 2010.03.22 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/event_function.asp"-->
<%
Dim evt_code 
	evt_code		= requestCheckVar(Request("evt_code"),10) 	
%>
<script language="javascript">
	
	//-- jsPopCal : �޷� �˾� --//
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	
	function jsWinnerSubmit(frm,v){
			
		if(!frm.evt_ranking.value){
			alert("����� �Է����ּ���");
			frm.evt_ranking.focus();
			return;
		}
		
		if(!IsDigit(frm.evt_ranking.value)){
			alert("����� ���ڸ� �Է°����մϴ�.");
			frm.evt_ranking.focus();
			return;
		}
		
		if(v == 'userseq'){
			if(!frm.evt_winner_seq.value){
				alert("��÷�ڸ� �Է����ּ���");
				frm.evt_winner_seq.focus();
				return;
			}
		}else{
			if(!frm.evt_winner_user.value){
				alert("��÷�ڸ� �Է����ּ���");
				frm.evt_winner_user.focus();
				return;			
			}		
		}
						
		if(confirm("����Ͻ� ������ ���� �Ǵ� ������ �Ұ����ϸ� ������ �ٷ� ����˴ϴ�.\n\n��� �Ͻðڽ��ϱ�? ")){
			frm.smode.value = v;
			frm.submit();
		}
	}
	    
    //����ǰ ���� ���
	function jsSetGiftKind(){
		var winkind;
		winkind = window.open('/admin/offshop/gift/popgiftKindReg.asp?giftkind_name='+document.frmWin.giftkind_name.value,'popkind','width=600, height=600;');
		winkind.focus();
	}

</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷�� ���</div>
<table width="100%" border=0 align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmWin" method="post" action="eventprize_process.asp">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="prize_add">
<input type="hidden" name="smode">
<tr>
	<td width="100" align="center" bgcolor="FFFFFF" colspan=3>�⺻����</td>
</tr>
<tr>
	<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td bgcolor="#FFFFFF">
		<%sbGetOptCommonCodeArr_off "evtprize_type", "", False,True,""%>
	</td>
</tr>		
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">���</td>
	<td bgcolor="#FFFFFF"><input type="text" size="2" name="evt_ranking"> ������ 0</td>
</tr>	
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����Ī<br>(Ƽ�ϸ���)</td>
	<td bgcolor="#FFFFFF"><input type="text" name="evt_rankname" size="20" maxlength="32"></td>
</tr>	
<% if evt_code = "" then %>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷Ȯ�αⰣ</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="evtprize_startdate" value="<%= left(now(),10) %>"  size="10" maxlength="10" onClick="jsPopCal('evtprize_startdate');" style="cursor:hand;">
		~<input type="text" name="evtprize_enddate" size="10"  maxlength="10" value="<%=dateadd("d",14,date())%>" onClick="jsPopCal('evtprize_enddate');" style="cursor:hand;">
	</td>
</tr>	
<% end if %>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷��<br>(�������ΰ���ȣ)</td>
	<td bgcolor="#FFFFFF">
		�޸ӷ� ����, ������� ( EX: 30,111 )<br>
		<textarea name="evt_winner_seq" rows="2" cols="60"></textarea>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="button" onclick="jsWinnerSubmit(frmWin,'userseq');" value="�������ΰ���ȣ�� ����" class="button">		
	</td>
</tr>
</table>
<br>
<table width="100%" border="0" align="center" class="a" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��÷��<br>(�̸����� ����)</td>
	<td bgcolor="#FFFFFF">
		�޸ӷ� ����, ������� ( EX: ������,�ѿ�� )<br>
		<textarea name="evt_winner_user" rows="2" cols="60"></textarea>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right" height="40">
		<input type="button" onclick="jsWinnerSubmit(frmWin,'username');" value="�̸����� ����" class="button">	
	</td>
</tr>
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->


			

	
				