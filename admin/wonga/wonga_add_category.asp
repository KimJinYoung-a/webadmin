<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ �׷� ī�װ� ���
' History : �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/wonga/wonga_month_class.asp"-->

<% 
dim menupos,gubun
	menupos = request("menupos")
	gubun = request("gubunbox")		'�����׷찪�� ������� ������ ���Ͽ� ���а��� �޾� �´�.

dim owongamonth_re,i,tmp
set owongamonth_re = new Cwongalist
	owongamonth_re.frectgubun = Request("gubunbox")
	owongamonth_re.fwongamonth_add()	
%>	
<%	
'###########################################################	�׷�� ����Ʈ�ڽ�
Sub DrawUserGubun(gubunbox,gubunid)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str
	
	'����� �˻� �ɼ� ���� DB���� ��������
	userquery = "select groupname from"
	userquery = userquery & " db_datamart.dbo.tbl_month_wonga_category"
	userquery = userquery & " group by groupname"
	userquery = userquery & " order by groupname asc"
	db3_rsget.Open userquery, db3_dbget, 1
	'response.write userquery&"<br>"
	
	response.write "<select onChange=javascript:check_gubun(this); name='" & gubunbox & "' "  '�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	if gubunid <> "" then					'���а��� ������ ������ ���ϵ��� disabled
		response.write "disabled"
	end if	
	response.write ">"		
	response.write "<option value=''"							'�ɼ��� ���� ������
		if gubunid ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">������뱸�� ����</option>"								'�����̶� �ܾ ��������.

	if not db3_rsget.EOF then
		do until db3_rsget.EOF
			if Lcase(gubunid) = Lcase(db3_rsget("groupname")) then 	'�˻��� �̸��� db�� ����� �̸��� ���ؼ� �´ٸ�, //
				tem_str = " selected"								'// �˻���� ����
			end if
			response.write "<option value='" & db3_rsget("groupname") & "' " & tem_str & ">" & db2html(db3_rsget("groupname")) & "</option>"
			tem_str = ""				'db3_rsget�� gubunid �����ϰ� �˻��� ������ ����
			db3_rsget.movenext
		loop
	end if
	response.write "</select>"
db3_rsget.close		
End Sub
%>

<script language="javascript">

<!-- ���� �˻�����-->
function check_gubun(frm)
{
	document.frmreg.groupname.value = "";
	document.frmreg.groupname.value = document.frmreg.gubunbox.value;
	document.frmreg.submit();
}
<!-- ���� �˻� ��-->

function form_submit(){
	if (document.frmreg.groupname.value=="")
	{
		alert('������ �Է��ϼ���');
		document.frmreg.groupname.focus();		 
	}
	else if (document.frmreg.category_box_0.value=="")
	{
		alert('ī�װ�0 ����  �Է��ϼ���');
		document.frmreg.category_box_0.focus();		 
	}
		else if (document.frmreg.field_box_0.value=="")
	{
		alert('ī�װ�1 �ʵ�1�� ����� �̸��� �Է��ϼ���.');
		document.frmreg.field_box_0.focus();		 
	}
	else
	{
		frmreg.action = "/admin/wonga/wonga_add_category_process.asp";
		frmreg.submit();
	}	
}

<!--�ʵ� �߰� ����-->
	function addAuthItem(frmreg)
	{		
		// ���߰�
		var oRow0 = tbl_auth.insertRow();
		oRow0.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};
		var oRow1 = tbl_auth.insertRow();
		oRow1.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};
		
		// ���߰� 
		var oCell1 = oRow0.insertCell();		
		var oCell2 = oRow1.insertCell();
		
	oCell1.innerHTML = "�ʵ��0 : <input type='text' name='field_box_0' size='30' maxlength='30'> &nbsp; ���ذ� : <input type='text' name='gijun_box_0' size='20' maxlength='20'> ���ڷ��Է��ϼ���"
		
	}
<!--�ʵ� �߰� ��-->
	
<!--�ʵ� �߰� ����
	function addAuthItem_field(frmreg)
	{		
		// ���߰�
		var oRows0 = tbl_auth.insertRow();
		oRows0.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};
		// ���߰� 
		var oCells1 = oRows0.insertCell();		
		
	oCells1.innerHTML = "�ʵ��0 : <input type='text' name='field_box_0' size='20' maxlength='20' value=''>"
	frmreg.flag_field.value = frmreg.flag_field.value+frmreg.flag.value+","
	
	}
�ʵ� �߰� ��-->
	
</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�׷� , ī�װ� ���</strong> / ���ذ��̶�? �̷���� �ϴ� ��ǥ �޼�ġ�� ���մϴ�.</font>
			</td>			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td><br></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!--ǥ ��峡-->

<% if owongamonth_re.ftotalcount = 0 then 
'##################################################################################################################	�׷� ���� ����	
%>

<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#ffffff><form name="frmreg" method="post" action="">
		<td align="center">
			��뱸��(�ʼ�) : 
		</td>
		<td colspan="3">
			<% DrawUserGubun "gubunbox", gubun %> &nbsp;&nbsp;&nbsp;<input type="hidden" name="gubun_submit" value="<%= gubun %>">
			�����Է� : <input type="text" name="groupname" size="20" maxlength="20"> (ex: ����)
		</td>
	</tr>
</table><br>
<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">	
	<tr bgcolor=ffffff>
		<td align="center">
			ī�װ�1�̸� : 
		</td>
		<td>
			<input type="text" name="category_box_0"> &nbsp; <input type="button" name="aa" value="�ʵ��߰�" onclick="addAuthItem(frmreg)">
		</td>
	</tr>
	<tr bgcolor=ffffff>
		<td colspan="4">
			<table width="100%" name="tbl_auth" id="tbl_auth" class="a" cellpadding="1" cellspacing="1"></table>
		</td>
		
	</tr>
</form>
</table>


<%
'##################################################################################################################	�����׷� ����
else %>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#ffffff><form name="frmreg" method="post" action="">
		<td align="center">
			��뱸��(�ʼ�) : 
		</td>
		<td colspan="3">
			<% DrawUserGubun "gubunbox", gubun %> &nbsp;&nbsp;&nbsp;<input type="hidden" name="gubun_submit" value="<%= gubun %>">
			�׷�� : <%= gubun %>
		</td>
	</tr>
	<%
	dim sql ,ftotalcount
		sql = "select"
		sql = sql & " category"
		sql = sql & " from db_datamart.dbo.tbl_month_wonga_category"
		sql = sql & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y'"
		sql = sql & " group by category" 	
	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"	
	ftotalcount = db3_rsget.recordcount
	db3_rsget.close
	%>	
	
	<% for i = 0 to ftotalcount - 1 %>
		<tr bgcolor=ffffff>
			<td align="center">
				ī�װ���<%= i %> (�ʼ�) : 
			</td>
			<td colspan="3"> <%= frectcategoryname(i,0) %><input type="hidden" name="groupname" size="20" maxlength="20" value="<%= gubun %>"></td>
		</tr>
		<%
		dim sql1 ,ffieldcount ,t
		sql1 = "select field"
		sql1 = sql1 & " from db_datamart.dbo.tbl_month_wonga_category"
		sql1 = sql1 & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y' and category='"& i &"'"
	
		db3_rsget.open sql1,db3_dbget,1
		'response.write sql1&"<br>"	
		ffieldcount = db3_rsget.recordcount
		db3_rsget.close
		%>
		<% for t = 0 to ffieldcount -1 %>
			<tr bgcolor=ffffff>
				<td align="center">�ʵ��:</td>
				<td><%= frectfieldname(i,t) %></td>
				<td align="center">���ذ�:</td>
				<td><%= frectgijunvalue(i,t) %></td>
			</tr>			
		<% next %>
	<% next %>
</table><br>
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=ffffff>
		<td align="center">
			<input type="hidden" name="add_category" value="<%= ftotalcount %>">�߰� ī�װ��� : 
		</td>
		<td colspan="3">
			<input type="text" name="category_box_0"> &nbsp; <input type="button" name="aa" value="�ʵ��߰�" onclick="addAuthItem(frmreg)">
		</td>
	</tr>
	<tr bgcolor=ffffff>
		<td colspan="4">
			<table width="100%" name="tbl_auth" id="tbl_auth" class="a" cellpadding="1" cellspacing="1"></table>
		</td>
		
	</tr>
</form>
</table>
<% end if %>
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right"><br><input type="button" value="�����ϱ�" onclick="form_submit();">&nbsp;
        	<input type="button" value="�ݱ�" onclick="javascript:window.close();"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->