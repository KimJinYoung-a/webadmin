<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_eventitem_group.asp
' Description :  �̺�Ʈ �׷���
' History : 2007.02.22 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
Dim eCode : eCode = Request("eC")
Dim idx : idx = Request("idx") 
dim arrList, cEGroup, intg

 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode
 	cEGroup.FGDisp = 1
  	arrList = cEGroup.fnGetEventItemGroup		
 set cEGroup = nothing
%> 
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
<!--
function TnGroupCodeSelect(gCode){
    $("#groupcode<%=idx%>",opener.document).val(gCode);
    self.close();
} 
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ �׷� �ڵ� ����</div><hr>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
 <tr>
 	<td>
        <%IF isArray(arrList) THEN %>
 		<select name="linkkind" onChange="TnGroupCodeSelect(this.value);">
            <option value="">����</option>
            <%FOR intg = 0 To UBound(arrList,2)%>
            <option value="<%=arrList(0,intg)%>"><%=db2html(arrList(1,intg))%></option>
            <%NEXT%>
        </select>
        <% else %>
        ��ϵ� �׷��� �����ϴ�. �׷� ������ ���� ���ּ���.
        <%END IF%>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->