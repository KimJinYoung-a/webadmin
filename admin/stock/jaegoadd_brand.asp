<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľ� �귣�庰 �������
' History : 2007�� 10�� 31�� �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim makerid , i
	makerid = request("makerid")		'�귣��� �˻��� ���� ����
	if makerid = "" then
		makerid = "����"
	end if 	
	
dim oip						'Ŭ��������
	set oip = new Cfitemlist		'������ ��Ż�� �ֱ�
	oip.frectmakerid = makerid
	oip.fbrandinsert()	
%>

<script language="javascript">
	function AnSelectAllFrame(bool){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.disabled!=true){
					frm.cksel.checked = bool;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}	
		
	function AnCheckClick(e){
		if (e.checked)
			hL(e);
		else
			dL(e);
	}	
	
	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				pass = ((pass)||(frm.cksel.checked));
			}
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
	
	function reg(upfrm){
	if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.drawitemid.value = upfrm.drawitemid.value + frm.itemid.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.drawitemid.value;
			var aa;
			aa = window.open("jaegoadd_brand_process.asp?drawitemid=" +tot, "reg","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
</script>	
<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get">
	<input type="hidden" name="drawitemid">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�귣�庰 ����ľ� �������</strong> / ���ϻ�ǰ�� ����ľ����ϰ�� ��ϵ��� �ʽ��ϴ�. </font>
			</td>		
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td><br>�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>&nbsp;
			<input type=button value="�˻�" onclick="document.frm.submit();">
			<br><br>
			<% if oip.ftotalcount > 0 then %>
				<input type="button" value="���" onclick="javascript:reg(frm);">
			<% end if %>
			<font color="red">�ѻ�ǰ�� �ɼ��� ������ ���� �Ұ�� ���� �Ѱ��� ���� �ϼ���.</font>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
	</form>
</table>
<!--ǥ ��峡-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<% if oip.ftotalcount > 0 then %>	 <!--���ڵ� ���� 0���� ũ�� -->
    <tr align="center" bgcolor="#DDDDFF">
   		<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td>�̹���</td>
		<td>��ǰ�ڵ�</td>
		<td>�귣��id</td>
		<td>��ǰ��</td>
		<td>�ɼ��ڵ�</td>
		<td>�ɼǸ�</td>	
		<td>����ľǿ����</td>
		</tr>

	<% for i=0 to oip.ftotalcount - 1 %>
		<form action="jaegoadd_brand_process.asp" name="frmBuyPrc<%=i%>" method="get">			<!--for�� �ȿ��� i ���� ������ ����-->
		<tr bgcolor="#FFFFFF">
			<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>	
			<td><img src="<%= oip.flist(i).fsmallimage %>" width=50 height=50><input type="hidden" name="smallimage" value="<%= oip.flist(i).fsmallimage %>"></td>	<!--'�̹��� -->
			<td><%= oip.flist(i).fitemid %><input type="hidden" name="itemid" value="<%= oip.flist(i).fitemid %>"></td>				 					<!--'��ǰ��ȣ	 -->
			<td><%= oip.flist(i).fmakerid %><input type="hidden" name="makerid" value="<%= oip.flist(i).fmakerid %>"></td>									 <!--'�귣��id -->
			<td><%= oip.flist(i).fitemname %><input type="hidden" name="itemname" value="<%= oip.flist(i).fitemname %>"></td>									 <!--'��ǰ�� -->
			<td><%= oip.flist(i).fitemoption %><input type="hidden" name="itemoption" value="<%= oip.flist(i).fitemoption %>"></td>							 <!--'�ɼ��ڵ� -->
			<td><%= oip.flist(i).fitemoptionname %><input type="hidden" name="itemoptionname" value="<%= oip.flist(i).fitemoptionname %>"></td>				 <!--'�ɼǸ� -->													
			<td><%= oip.flist(i).fbasicstock %><input type="hidden" name="basicstock" value="<%= oip.flist(i).fbasicstock %>"></td>								 <!--'����ľǻ��� -->
		</tr>
	    </form>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="left">
        <% if oip.ftotalcount > 0 then %>
			<input type="button" value="���" onclick="javascript:reg(frm);">
		<% end if %>
        <input type="button" value="�ݱ�" onclick="javascript:window.close();">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/lib/db/dbclose.asp" -->