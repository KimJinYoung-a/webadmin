<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/photo_req/boardCls.asp"-->
<%
'####################################################
' Description :  �۾��� ��� �� ����Ʈ ������
' History : 2012.03.09 ������ ����
'####################################################
Dim arrList,i
Dim utype
Dim userid, sCodeDesc, uno, isusing
Dim worker, sMode, selectBoxName

uno		= requestCheckVar(Request("uno"),10)
userid  = requestCheckVar(Request("uid"),10)
utype =  requestCheckVar(Request("seltype"),10)

IF utype = "" THEN utype = "1"
Set worker = new CCoopUserList
IF uno <> "" THEN
	worker.FMode = "U"
	worker.FUser_no = uno
Else
	worker.FMode = "I"
	worker.FCodeType = utype
END IF
worker.fnGetCoopUserList
%>
<script language="javascript">
<!--
	// �ڵ�Ÿ�� �����̵�
	function jsSetCode(no,stype){
		self.location.href = "PopUserList.asp?uno="+no+"&seltype="+stype;
	}

	//�ڵ� �˻�
	function jsSearch(){
		document.frmSearch.submit();
	}

	//�ڵ� ���
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.uid.value) {
			alert("�۾���ID�� �Է��� �ּ���");
			frm.uid.focus();
			return false;
		}

		if(!frm.uname.value) {
			alert("�۾��ڸ��� �Է��� �ּ���");
			frm.uname.focus();
			return false;
		}

		return true;
	}

	//ID �˻� �˾�â
	function jsSearchID(frmName,compName,userName){
	    var compVal = "";
	    try{
	        compVal = eval("document.all." + frmName + "." + compName + "." + userName).value;
	    }catch(e){
	        compVal = "";
	    }

	    var popwin = window.open("/admin/photo_req/popUserSearch.asp?frmName=" + frmName + "&compName=" + compName + "&userName=" + userName + "&rect=" + compVal,"popBrandSearch","width=800 height=400 scrollbars=yes resizable=yes");

		popwin.focus();
	}
//-->
</script>

<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td colspan="2"><!--//�ڵ� ��� �� ����-->
		<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a" >
		<form name="frmReg" method="post" action="procUser.asp" onSubmit="return jsRegCode();">
		<input type="hidden" name="sM" value="<%=worker.FMode%>">
		<tr>
			<td>	+ �۾��� ��� �� ����</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�۾���Ÿ��</td>
					<td bgcolor="#FFFFFF">
						<select name="seltype">
							<option value="">-����-</option>
							<option value="1" <%IF Cstr(utype)="1" THEN%>selected<%END IF%>>����׷���</option>
							<option value="2" <%IF Cstr(utype)="2" THEN%>selected<%END IF%>>��Ÿ�ϸ���Ʈ</option>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"  width="100" align="center">�۾���ID</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="10" maxlength="10" name="uid" <% If worker.FMode = "U" Then response.write "value="&worker.FUserList(0).FUser_id&"" End If%> readonly>
						<input type="button" class="button" value="ID�˻�" onclick="jsSearchID('frmReg','uid','uname');" >
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">�۾��ڸ�</td>
					<td bgcolor="#FFFFFF">
						<input type="text" size="10" maxlength="16" name="uname" <% If worker.FMode = "U" Then response.write "value="&worker.FUserList(0).FUser_name&"" End If%> readonly>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EFEFEF"   align="center">��뿩��</td>
					<td bgcolor="#FFFFFF">
					<% IF worker.FMode = "U" Then %>
						<input type="radio" value="Y" name="isusing" <% If worker.FUserList(0).FUser_useyn = "Y" Then response.write "checked" End If%> >���
						<input type="radio" value="N" name="isusing" <% If worker.FUserList(0).FUser_useyn = "N" Then response.write "checked" End If%> >������
					<% Else %>
						<input type="radio" value="Y" name="isusing" checked >���
						<input type="radio" value="N" name="isusing">������
					<% End If %>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="right"><input type="image" src="/images/icon_save.gif">
				<a href="javascript:jsSetCode('','')"><img src="/images/icon_cancel.gif" border="0"></a>
			</td>
		</tr>
		<tr>
			<td colspan="2"><hr width="100%"></td>
		</tr>
		</form>
		</table>
	</td>
</tr>
</table>
<table width="385" border="0" cellpadding="3" cellspacing="0" class="a" >
<form name="frmSearch" method="post" action="PopUserList.asp">
<tr>
	<td colspan="2">+ �۾��ڸ���Ʈ</td>
</tr>
<tr>
	<td>�۾���Ÿ�� :
		<select name="seltype" onChange="jsSearch(this.value);">
			<option value="">-����-</option>
			<option value="1" <%IF Cstr(utype)="1" THEN%>selected<%END IF%>>����׷���</option>
			<option value="2" <%IF Cstr(utype)="2" THEN%>selected<%END IF%>>��Ÿ�ϸ���Ʈ</option>
		</select>
	</td>
	<td align="right"></td>
</tr>
<tr>
	<td colspan="2">
		<div id="divList" style="height:305px;overflow-y:scroll;">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
		<tr bgcolor="#EFEFEF">
			<td align="center">�۾���ID</td>
			<td align="center">�۾��ڸ�</td>
			<td align="center">��뿩��</td>
			<td align="center">ó��</td>
		</tr>
<%
	If worker.Fresultcount > 0 THEN
		For i = 0 To worker.fresultcount -1
%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=worker.FUserList(i).FUser_id%></td>
			<td align="center"><%=worker.FUserList(i).FUser_name%></td>
			<td align="center"><%=worker.FUserList(i).FUser_useyn%></td>
			<td align="center"><input type="button" value="����" onClick="javascript:jsSetCode('<%=worker.FUserList(i).FUser_no%>','<%=worker.FUserList(i).FUserType%>');" class="input_b"></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF"><td colspan="5" align="center">��ϵ� ������ �����ϴ�.</td></tr>
<%End if%>
		</table>
		</div>
	</td>
</tr>
</form>
</table>
<% Set worker = nothing%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->