<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ������ �������� ����
' History : 2011.09.23 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<%
Dim icomm_Cd  
Dim clsAccount, arrAccountData  
 	icomm_Cd	= requestCheckvar(Request("iCCd"),10)
 	
set clsAccount = new COpExpAccount 
	clsAccount.Fcomm_cd = icomm_Cd
	arrAccountData = clsAccount.fnGetAccountData
set clsAccount = nothing 
%>
<select name="selAI" id="selAI" class="select">
<option value="0">--����--</option>
<% sbOptAccount arrAccountData, 0%>
</select>