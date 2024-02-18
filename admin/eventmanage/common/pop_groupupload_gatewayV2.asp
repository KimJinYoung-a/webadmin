 <%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_eventitem_groupImage.asp
' Description :  �̺�Ʈ �׷� �̹��� ����
' History : 2007.02.22 ������ ����
'			2015.02.12 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls_V2.asp"--> 
<%
Dim eCode : eCode = Request("eC")
dim eChannel : eChannel = requestCheckVar(Request("eCh"),1)
 
dim cEGroup,arrGroup,vYear,intg

set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode
	cEGroup.FEChannel = eChannel
  	arrGroup = cEGroup.fnGetEventItemGroup
  	vYear = cEGroup.FRegdate
set cEGroup = nothing
	 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<div id="divGC">
 <%IF isArray(arrGroup) THEN %>
	<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
		<td>�׷��ڵ�</td>					
		<td>�����׷�</td>
		<td>�׷��</td>
		<td>���ļ���</td>					
		<td>�̹���</td>
		<td>���ÿ���</td>
		<td>����</td>
	</tr>
	<%
	dim sumi, i
	FOR intg = 0 To UBound(arrGroup,2)
	sumi = 0 
	%>				   						
	<tr <%if not arrGroup(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
		<td  align="center"><%IF arrGroup(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrGroup(0,intg)%>
		    <% if intg < UBound(arrGroup,2) and eChannel ="M" then 
					    for i = 1 to (UBound(arrGroup,2)-intg)%>
					    <%if arrGroup(9,intg) = arrGroup(9,intg+i) then
					        sumi = sumi + 1  
					         %>
					    + <%=arrGroup(0,intg+i)%>
					    
					    <%else 
					     exit for
					    end if 
					    next
					end if
					    %> 
		 </td>						
		<td  align="center"><%IF isnull(arrGroup(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrGroup(5,intg)%>]<%=db2html(arrGroup(7,intg))%><%END IF%></td>	
		<td  align="center"><%=db2html(arrGroup(1,intg))%></td>	
		<td  align="center"><%=arrGroup(2,intg)%></td>									   									
		<td  align="center">   
			<a href="javascript:jsImgView('<%=arrGroup(3,intg)%>');"><img src="<%=arrGroup(3,intg)%>" width="50" border="0"></a>  
		</td>		
		<td  align="center"><%if arrGroup(8,intg) then%>Y<%else%>N<%end if%>&nbsp; <input type="button" name="btnA" value="����" onclick="jsDispGroup('<%=arrGroup(0,intg)%>','<%if arrGroup(8,intg) then%>0<%else%>1<%END IF%>')"  class="button"></td>			   									
		<td  align="center">
			<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrGroup(0,intg)%>','<%=eChannel%>')" class="button">
			<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrGroup(0,intg)%>')"  class="button">-->
			<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrGroup(0,intg)%>','<%=eChannel%>')"  class="button">
			<% IF arrGroup(5,intg) = 0 THEN %>
			<%if eChannel = "M" then%>
			<% 		Response.Write "<a href='" & mobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			 <%else%>
			 <% 		Response.Write "<a href='" & wwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrGroup(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			 <%end if%>
			<% END IF %>
		</td>					   									
	</tr>
	<%
	  intg = intg+sumi
	NEXT%>
	</table>
<%END IF%>	
</div> 
<script type="text/javascript">   
	<%if eChannel ="M" then%>
	$("#divMFrm3", opener.document).html($("#divGC").html()); 
	<%else%>
	$("#divFrm3", opener.document).html($("#divGC").html());
  	<%end if%> 
	 window.close();
	 
 //	});
</script>