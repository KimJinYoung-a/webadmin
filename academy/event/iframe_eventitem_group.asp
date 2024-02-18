<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ �׷캸��
' History : 2010.09.28 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim eCode : eCode = RequestCheckvar(Request("eC"),10)
Dim sType : sType = RequestCheckvar(Request("T"),10)
Dim cEGroup, arrList,intg

set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode  	
	arrList = cEGroup.fnGetEventItemGroup		
set cEGroup = nothing 
%>

<script language="javascript" defer>

  function jsAddGroup(eCode,gCode){
  	var winG
  	
  	<% if eCode = "" then %>
  	alert('�̺�Ʈ ó�� ��Ͻÿ��� �׷����� ���� �ϽǼ� �����ϴ�. \n�̺�Ʈ�� �켱 �����Ͻ��� �������� �׷����� ������ �ּ���');
  	return;
  	<% end if %>
  	
  	winG = window.open('pop_eventitem_group.asp?eC='+eCode+'&eGC='+gCode,'popG','width=600, height=500');
  	winG.focus();
  }
  
  function jsDelGroup(eCode,gCode){
   if(confirm("������ �����׷�� ��� �����˴ϴ�. �����Ͻðڽ��ϱ�? ")){
  	document.frmD.eGC.value = gCode; 
  	document.frmD.submit();
  	}
  }
  
  //-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
	function jsImgView(sImgUrl){
	 var wImgView;
	 wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	 wImgView.focus();
	}
	

</script>

<%IF sType="1" THEN%><body onunload="opener.location.href='eventitem_regist.asp?eC=<%=eCode%>';"><%END IF%>
<form name="frmD" method="post" action="event_process.asp">
<input type="hidden" name="imod"  value="gD">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="">
</form>
<table width="650" border="0" cellpadding="1" cellspacing="3" class="a">		   							   					
<tr>
	<td>					   						
		<input type="button" value="�׷��߰�" onClick="jsAddGroup('<%=eCode%>','');" class="input_b">
	</td>
</tr>
<tr>
	<td>
		<%IF isArray(arrList) THEN %>
			<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center">
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>					
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">�����׷�</td>
				<td bgcolor="<%= adminColor("tabletop") %>">�׷��</td>
				<td width="50" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>					
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">�̹���</td>
				<td width="100" bgcolor="<%= adminColor("tabletop") %>">����</td>
			</tr>
			<%FOR intg = 0 To UBound(arrList,2)%>				   						
			<tr>
				<td  align="center" bgcolor="#FFFFFF"><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
				<td  align="center" bgcolor="#FFFFFF"><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
				<td  align="center" bgcolor="#FFFFFF"><%=db2html(arrList(1,intg))%></td>	
				<td  align="center" bgcolor="#FFFFFF"><%=arrList(2,intg)%></td>									   									
				<td  align="center" bgcolor="#FFFFFF"><%IF arrList(3,intg) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a><%END IF%></td>					   									
				<td  align="center" bgcolor="#FFFFFF">
					<input type="button" name="btnU" value="����" onclick="jsAddGroup('<%=eCode%>','<%=arrList(0,intg)%>')" class="button">
					<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
				</td>					   									
			</tr>
			<%NEXT%>
			</table>
		<%END IF%>	
	</td>
</tr>				
</table>
	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->