<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/iframe_eventitem_group.asp
' Description :  �̺�Ʈ �׷캸��
' History : 2007.02.22 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode : eCode = Request("eC")
Dim sType : sType = Request("T")
Dim ekind : ekind = Request("ekind")
Dim cEGroup, arrList,intg, vYear

 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode  	
 	cEGroup.FRectGroupDelInc ="N"
  	arrList = cEGroup.fnGetEventItemGroup		
  	vYear = cEGroup.FRegdate
 set cEGroup = nothing
 
dim isDelGrp 
%>
<script language="javascript" defer>
<!-- 
function jsAddGroup(eCode,gCode){
	var winG
	winG = window.open('pop_eventitem_group.asp?yr=<%=vYear%>&eC='+eCode+'&eGC='+gCode,'popG','width=600, height=500');
	winG.focus();
}

function jsDelGroup(eCode,gCode){
	if(confirm("������ �����׷�� ��� �����˴ϴ�. �����Ͻðڽ��ϱ�? ")){
		document.frmD.eGC.value = gCode; 
		document.frmD.submit();
	}
}

function popRegItem(eCode, gCode){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/event/eventitem_regist.asp?eC='+eCode+'&selG='+gCode,'pImg','width=900,height=800');
	wImgView.focus();
}

//-- jsImgView : �̹��� Ȯ��ȭ�� ��â���� �����ֱ� --//
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/eventmanage/common/pop_event_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}

//-->
</script>
<%IF sType="1" THEN%><body onunload="opener.location.href='eventitem_regist.asp?eC=<%=eCode%>';"><%END IF%>
<form name="frmD" method="post" action="event_process.asp">
<input type="hidden" name="imod"  value="gD">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eGC" value="">
</form>
<table width="800" border="0" cellpadding="1" cellspacing="3" class="a">		   							   				

	
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
				<% isDelGrp = (arrList(8,intg)="N")  %>
				<tr >
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%=db2html(arrList(1,intg))%></td>	
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%=arrList(2,intg)%></td>									   									
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> ><%IF arrList(3,intg) <> "" THEN%><a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a><%END IF%></td>					   									
					<td  align="center" <%= CHKIIF(isDelGrp,"bgcolor='#CCCCCC'","bgcolor='#FFFFFF'") %> >
					    <% if ( isDelGrp) then %>
					    ������ �׷�
					    <% else %>
						<input type="button" name="btnU" value="����" onclick="jsAddGroup('<%=eCode%>','<%=arrList(0,intg)%>')" class="button">
						<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
						<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">
						<% IF arrList(5,intg) = 0 THEN %>
						<%
							'�̺�Ʈ ������ ���� ����Ʈ��ũ ������ ����
							Select Case ekind
								Case "26"		'�����
									Response.Write "<a href='" & vmobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>�̸�����</a>"
								Case Else		'�������� �� ��Ÿ
									Response.Write "<a href='" & vwwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg)&" ' target='_blank'>�̸�����</a>"
							End Select
						%>
						<% END IF %>
					    <% end if %>
					</td>					   									
				</tr>
				<%NEXT%>
				</table>
			<%END IF%>	
		</td>
	</tr>				
</table>
	
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->