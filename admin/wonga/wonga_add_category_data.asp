<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������������ �׷� ī�װ� ����Ÿ ���
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
dim menupos,gubun,category_add
	menupos = request("menupos")
	gubun = request("gubunbox")		'�����׷찪�� ������� ������ ���Ͽ� ���а��� �޾� �´�.
	category_add = request("category_add_box")
	
	if category_add = "" then
		category_add = 1
	end if	
	
dim owongamonth_re,i
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
	else if (document.frmreg.yyyy.value=="")
	{
		alert('�⵵�� �Է��ϼ���');
		document.frmreg.yyyy.focus();		 
	}
	else if (document.frmreg.mm.value=="")
	{
		alert('���� �Է��ϼ���');
		document.frmreg.mm.focus();		 
	}
	else if (document.frmreg.count.value=="")
	{
		alert('�� ������ �Է��ϼ���');
		document.frmreg.count.focus();		 
	}
	else
	{
		frmreg.action = "/admin/wonga/wonga_add_category_data_process.asp";
		frmreg.submit();
	}	
}

<!-- �⵵�� �� �ߺ�üũ ����-->
function yyyymmcheck(){
	if (document.frmreg.groupname.value=="")
	{
		alert('������ �Է��ϼ���');
		document.frmreg.groupname.focus();		 
	}
	else if (document.frmreg.yyyy.value=="")
	{
		alert('�⵵�� �Է��ϼ���');
		document.frmreg.yyyy.focus();		 
	}
	else if (document.frmreg.mm.value=="")
	{
		alert('���� �Է��ϼ���');
		document.frmreg.mm.focus();		 
	}
	else
	{
		var yyyy = frmreg.yyyy.value;
		var mm = frmreg.mm.value;
		var groupname = frmreg.groupname.value;
		var popup = window.open('/admin/wonga/wonga_yyyymm_check.asp?yyyy='+yyyy+'&mm='+mm+'&groupname='+groupname,'yyyymmcheckpopup','width=1,height=1,scrollbars=yes,resizable=yes');
		popup.focus();	
	}	
}
<!-- �⵵�� �� �ߺ�üũ ��-->

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
			<font color="red"><strong>�׷� , ī�װ� ����Ÿ ���</strong> / ���ذ��̶�? �̷���� �ϴ� ��ǥ �޼�ġ�� ���մϴ�.</font>
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
			���ñ׷� : <input type="text" name="groupname" size="20" maxlength="20" disabled> (ex: ����)
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
		<td colspan="5">
			<% DrawUserGubun "gubunbox", gubun %> &nbsp;&nbsp;&nbsp;<input type="hidden" name="gubun_submit" value="<%= gubun %>">
			�����Է� : <input type="text" name="groupname" size="20" maxlength="20" value="<%= gubun %>" disabled> (ex: ����)
		</td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align="center">
			��,�� �Է�(�ʼ�) : 
		</td>
		<td colspan="5">
			<input type="text" name="yyyy" size="4" maxlength="4"> 
			<input type="text" name="mm" size="2" maxlength="2"> (ex: 2007 , 01)
			<input type="button" name="checkbutton" value="�ߺ�üũ(�ʼ�)" onclick="yyyymmcheck();">
		</td>
	</tr>
	<tr bgcolor=#ffffff>
		<td align="center">
			�� ����(�ʼ�) :  
		</td>
		<td colspan="5">
			<input type="text" name="count" size="20" maxlength="20"> ex: ���� �� ������ &nbsp;&nbsp;&nbsp;
			<font color="red">��갪 = ����(2007��01��)ī�װ�1(�ʵ�1)�� / �Ѽ���(���������)</font>
		</td>
	</tr>
</table>	
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
<br>	
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">	
	<% for i = 0 to ftotalcount - 1 %>
		<tr bgcolor=ffffff>
			<td align="center">
				ī�װ���<%= i %> (�ʼ�) : 
			</td>
			<td colspan="5"> <%= frectcategoryname(i,0) %>
			<input type="hidden" name="category_box_0" size="20" maxlength="20" value="<%= frectcategoryname(i,0) %>">
			<input type="hidden" name="groupname" size="20" maxlength="20" value="<%= gubun %>"></td>
			
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
				<td align="center">�ʵ�� : </td>
				<td><input type="hidden" name="field_box_0" size="20" maxlength="20" value="<%= frectfieldname(i,t) %>"> <%= frectfieldname(i,t) %></td>
				<td align="center">���ذ� : </td>
				<td><input type="text" name="gijun_box_0" size="20" maxlength="20" value="<%= frectgijunvalue(i,t) %>"></td>
				<td>�� : </td>
				<td><input type="text" name="value_box_0" size="20" maxlength="20" value=""></td>
			</tr>			
		<% next %>
	<% next %>
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
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->