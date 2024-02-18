<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Response.CharSet = "euc-kr"
	
	Dim vUserID, vEmpNo
	vEmpNo = requestCheckvar(Request("empno"),32)
	
	If vEmpNo = "" Then
		dbget.close
		Response.End
	End If

	Dim vStaffImage, vStaffName, vStaffID, vStaffPartName, vStaffPosit, vStaffJob, vStaffEmail, vStaffHP, vStaffPhone, vStaffDirect, vStaffExt, vStaffMyWork
	Call sbOrganizationChartOne(vEmpNo)
	
%>
<div class="panel2 pad20 tMar20">
	<div class="ftLt col11 ct">
		<p style="width:124px; border:2px solid #fff; margin:0 auto;"><img src="<%=CHKIIF(vStaffImage="","http://webadmin.10x10.co.kr/images/partner/profile_defaultimg.png",vStaffImage)%>" alt="<%=vStaffName%> ����" style="width:120px"/></p>
	</div>
	<div class="ftRt" style="width:80%;">
		<ul class="listLine">
			<li><strong>�̸� (���̵�)</strong><span><%=vStaffName%> (<%=vStaffID%>)</span></li>
			<li><strong>�μ�</strong><span><%=vStaffPartName%></span></li>
			<% if C_ADMIN_AUTH or C_PSMngPart or C_MngPart then %><li><strong>���� (��å)</strong><span><%=vStaffPosit%> <%=CHKIIF(isNull(vStaffJob),"","("&vStaffJob&")")%></span></li><% end if %>
			<li><strong>E-mail</strong><span><a href="mailto:<%=vStaffEmail%>"><%=vStaffEmail%></a></span></li>
			<li><strong>�޴���ȭ��ȣ</strong><span><%=vStaffHP%></span></li>
			<li><strong>ȸ����ȭ</strong><span><%=vStaffPhone%> / ���� : <%=vStaffDirect%> <%=CHKIIF(vStaffExt="","","("&vStaffExt&")")%></span></li>
			<li><strong>������</strong><span><%=vStaffMyWork%></span></li>
		</ul>
	</div>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->