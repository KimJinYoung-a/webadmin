<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��÷�� ���
' History : 2008.04.11 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'--------------------------------------------------------
' ��������
'--------------------------------------------------------
Dim iDFSeq
Dim clsDF
Dim iDFType, sTitle, dPrizeDate,sDFTypeDesc
Dim eCode,egKindCode
 
 iDFSeq  = requestCheckVar(request("iDFS"),10)
 eCode = 1				'�������ΰŽ� �̺�Ʈ ��ȣ
 egKindCode = iDFSeq	'�������ΰŽ�ȸ��
'--------------------------------------------------------
' �̺�Ʈ ������ ��������
'--------------------------------------------------------
	 set clsDF = new CDesignFingers
	clsDF.FDFSeq = iDFSeq		
	clsDF.fnGetDFSummary	
	
	iDFType 	 = clsDF.FDFType 	
	sTitle  	 = clsDF.FTitle  		
	dPrizeDate   = clsDF.FPrizeDate 
	sDFTypeDesc  = clsDF.FDFTypeDesc
	set clsDF = nothing
	
	
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td>		
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> ��÷���� : �ѹ� ��ϵ� ��÷�ڴ� ����� �� �����ϴ�. �Է½� ������ �ּ���</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">�ΰŽ�ID</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=iDFSeq%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDFTypeDesc%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sTitle%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷��ǥ��</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=dPrizeDate%></td>					 
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- ��÷�� ���-->		
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
		<!-- /��÷�� ���-->
	</td>
</tr>	
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->