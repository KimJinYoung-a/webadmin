<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : ī�װ� ����Ʈ
' History : 2012.08.07 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp" -->
<%
Dim ipcateidx, icateidx
Dim clsAcc, arrList 
Dim sVar
  sVar			= requestCheckvar(Request("sVar"),5)
 	ipcateidx	= requestCheckvar(Request("selCL"),10)
 	icateidx	= requestCheckvar(Request("selC"),10) 
Set clsAcc = new CAccCategory 
 
%>
<select name="<%=sVar%>"	id="<%=sVar%>" class="select">
<option value="0">--����--</option>
<% 	IF ipcateidx > 0 THEN
	clsAcc.sbGetOptAccCategory 2,ipcateidx,icateidx 
	END IF%>
</select> 
<%Set clsAcc = nothing%>
<!-- #include virtual="/lib/db/dbclose.asp" -->