<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ������û�� ���
' History : 2011.03.14 ������  ����
' 0 ��û/1 ������/ 5 �ݷ�/7 ����/ 9 �Ϸ�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->  
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payManagerCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"--> 
<script language='javascript'>
function jsAttachDoc(v1,v2){
    var iURI = "popAddDoc.asp?idx="+v1;
    var popwin = window.open(iURI,'popAddDoc','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popTmsBaCUST(v1){
    var iURI = "/admin/approval/comm/popTmsBaCust.asp?cust_cd="+v1;
    var popwin = window.open(iURI,'popTmsBaCUST','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<table width="100%" cellpadding="5" cellspacing="1" class="a"  style="padding-bottom:50px;" >  
<tr>
	<td>
		<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a"   border="0" >
        <tr>
			<td>
				<table width="100%" align="left" cellpadding="5" cellspacing="1" class="a" border="0" bgcolor=#BABABA>
				<tr bgcolor="#E6E6E6" align="center">
					<td rowspan="2" width="80">÷�μ���</td>
					<td width="120">��������</td>
					<td>���ó���</td>
					<td><input type="button" value="�߰�" onClick="jsAttachDoc('idx','ret')"></td>
				</tr>
				<tr  bgcolor="#FFFFFF">
					<td align="center" valign="top" width="120">
						 <select name="paydoctype">
						 <option value="">����
						 <option value="1">���ݰ�꼭
						 <option value="2">���ݿ�����
						 <option value="3">(��Ÿ)������
						 <option value="11">����ڵ���� �纻
                         <option value="12">����纻
                         <option value="21">�ŷ���ǥ
                         <option value="99">��Ÿ����
						 </select>
					</td>
					<td> 
					    <input type="button" value="�ŷ�ó����" onclick="popTmsBaCUST('');">
					    
					    <!--
						<div id="dFile"> 
						</div>
						<input type="text" name="sL" size="60" maxlength="120"><input type="button" value="����÷��" class="button" onClick="jsAttachFile();"> 
                        -->						
					</td>
					<td> 
					    ����
				    </td>
				</tr>
				</table>
			</td>
		</tr>
</table> 
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" --> 
