<%@ language=vbscript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<%
'########################################################### 
' Description :  ���δ���ǰ �󼼸���Ʈ ��������
' History : 2014.01.06 ������ ����
'						currstate: 0-���ιݷ�,1-���δ��,2-���κ���,5-���δ��(���û),7-���οϷ�,9-��ü���
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"--> 
<%
Dim clsWait, itemid ,arrlist, intLoop
itemid =  requestCheckvar(Request("itemid"),16)
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog
 set clsWait = nothing
%>
 <%IF isArray(arrList) THEN %>
<table border=0 cellpadding=5 cellspacing=5 bgcolor="#EFEFEF" class="a"> 
	<tr>
		<td>�ӽ��ڵ�: [<%=itemid%>]<hr width="100%">
			<%
				For intLoop = 0 To UBound(arrList,2)
			%>
			<font color="<%=GetCurrStateColor(arrList(2,intLoop))%>"><%=fnGetCurrStateShortName(arrList(2,intLoop))%></font>: <%=arrList(4,intLoop)%><br>
			<%		
				Next
			%> 
		</td>
	</tr>
</table>
<%END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->