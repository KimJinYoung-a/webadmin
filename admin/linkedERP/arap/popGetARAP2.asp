<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����׸� ����Ʈ - ������
' History : 2011.11.15 ������  ����
'	jsSetARAP ��ũ��Ʈ �Լ� opener���� �����ؼ� ����ó��
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/arapCls.asp"-->
<%
Dim isUseSerp : isUseSerp = false

Dim clsARAP
Dim arrList, intLoop, taxKey
Dim sARAP_GB,sCASH_FLOW,sARAP_NM, sAcc

sARAP_GB = requestCheckvar(Request("rdoGB"),3)
sCASH_FLOW = requestCheckvar(Request("selFlow"),3)
sARAP_NM = requestCheckvar(Request("sNM"),50)
sAcc		=   requestCheckvar(Request("sAC"),50)
taxKey	= request("taxKey")

Set clsARAP = new CARAP
	 clsARAP.FARAP_GB		=sARAP_GB
	 clsARAP.FCASH_FLOW =sCASH_FLOW
	 clsARAP.FARAP_NM   =sARAP_NM
	 clsARAP.FACC				= sACC
	arrList = clsARAP.fnGetARAPCD
Set clsARAP = nothing
%>
<!-- #include virtual="/lib/db/dbClose.asp" -->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="javascript">
    //alert('2016/04/30 sERP ���׷��̵� �۾����Դϴ�. ������� ������. ������ ���� ���.');
	function jsGetErp(){
		location.href = "procGetErp.asp";
	}

	function jsGetArapInfo(){
		var winInfo = window.open("/admin/approval/eapp/popArapInfo.asp","popInfo",'width=1024, height=900, scrollbars=yes,resizable=yes');
		winInfo.focus();
	}
	function chromeOpenerFuncBug(a, b, c, d){
		window.opener.document.frmAct.mode.value = "modiArapCD"
		window.opener.document.frmAct.arap_cd.value = a;
		window.opener.document.frmAct.taxKey.value = "<%= taxKey %>";
		window.opener.document.frmAct.matchSeq.value="0"
		window.opener.document.frmAct.submit();
		self.close();
	}
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
	<td><table class="a" width="100%">
		<tr>
			<td width="600"><strong>�����׸�  ����</strong></td>
		<%IF C_MngPart OR C_ADMIN_AUTH or C_PSMngPart THEN%><td align="right"><input type="button" class="button" value="ERP��ϼ���" onClick="jsGetErp();"></td><%END IF%>
			<td align="right">&nbsp;<input type="button" class="button" value="�����׸�з�ǥ" onClick="jsGetArapInfo();"></td>
		</tr>
		<tr>
			<td colspan="3">  <hr width="100%"></td>
		</tr>
	</table>
</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popGetARAP2.asp">
			<input type="hidden" name="taxKey" value="<%= taxKey %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
				<td align="left">
					����:
					<input type="radio" name="rdoGB" value=""<%IF sARAP_GB="" THEN%>checked<%END IF%>>��ü
					<input type="radio" name="rdoGB" value="1" <%IF sARAP_GB="1" THEN%>checked<%END IF%>>����
					<input type="radio" name="rdoGB" value="2" <%IF sARAP_GB="2" THEN%>checked<%END IF%>>����
					&nbsp; &nbsp; &nbsp;
					�з�:
					<select name="selFlow">
						<option value="">��ü</option>
						<option value="001"  <%IF sCASH_FLOW="001" THEN%>selected<%END IF%>>����</option>
						<option value="002"  <%IF sCASH_FLOW="002" THEN%>selected<%END IF%>>����</option>
						<option value="003"  <%IF sCASH_FLOW="003" THEN%>selected<%END IF%>>�繫</option>
					</select>
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>�����׸��: <input type="text" name="sNM" value="<%=sARAP_NM%>" size="20">
					&nbsp;�����������: <input type="text" name="sAC" value="<%=sACC%>" size="20">
				</td>
			</tr>
		</form>
		</table>
	</td>
</tr>
<% if (C_ERP_VERSION <> "") then %>
<tr>
	<td>*���ؿ����� <%= Year(Now()) %>�� �Դϴ�.</td>
</tr>
<% end if %>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td>�ڵ�</td>
		 	<td>����</td>
			<td>�з�</td>
			<td>�����׸�</td>
			<td>�����������</td>
			<% if (NOT isUseSerp) then %><td>����/����ŷ�����</td><% end if %>
			<td>����</td>
		</tr>
		<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=arrList(0,intLoop)%></td>
		 	<td><%=fnGetARAP_GB(arrList(1,intLoop))%></td>
		 	<td><%=fnGetARAP_Cash(arrList(3,intLoop))%></td>
		 	<td><%=arrList(2,intLoop)%></td>
		 	<td align="left">[<%=arrList(9,intLoop)%>] <%=arrList(5,intLoop)%></td>
		 	<% if (NOT isUseSerp) then %><td><%=arrList(7,intLoop)%></td><% end if %>
			<td><input type="button" class="button" value="����" onClick="chromeOpenerFuncBug('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>');"> </td>
		</tr>
	<%	Next
		END IF%>
		</table>
	</td>
</tr>
</table>
<!-- ������ �� -->
</body>
</html>
