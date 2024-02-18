<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/v2/pop_eventgroup_selCopy.asp
' Description :  이벤트  상품 복사 그룹 선택
' History : 2015.04.22 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<%
Dim eCode : eCode = requestCheckVar(Request("eC"),10) 
Dim itemidArr : itemidarr = Request("itemidarr")
dim eChannel : eChannel = requestCheckVar(Request("eCh"),1)
Dim cEGroup, arrGroup ,intLoop
 
    set cEGroup = new ClsEventGroup
 		cEGroup.FECode = eCode  
 		cEGroup.FEChannel = eChannel
  		arrGroup = cEGroup.fnGetEventItemGroup	 
 	set cEGroup = nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript">
 
</script>

<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 이벤트 복사할 그룹 선택 </div><hr> 
<form name="frmG" method="post" action="eventgroup_process.asp"   onSubmit="return jsGroupSubmit(this);">
		<input type="hidden" name="eC" value="<%=eCode%>">
		<input type="hidden" name="eGC" value="<%=eGCode%>">
		<input type="hidden" name="mode" value="<%=sM%>">  
		<input type="hidden" name="eCh" value="<%=eChannel%>">  
		<input type="hidden" name="itemidArr" value="<%=itemidArr%>">
<table width="580" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
 <tr>
 	<td>
 		
			<select name="eG" class="select">
          	<%IF isArray(arrGroup) THEN
          		For intLoop = 0 To UBound(arrGroup,2)
          	%>
          		<option value=" <%=arrGroup(0,intLoop)%>" ><%IF arrGroup(5,intLoop) <> 0 THEN%>└&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)</option>
          <%	Next 
            ELSE	
          %>
          <option value=""> --그룹없음--</option>
          <%END IF%>	
          	</select>
	</td>
</tr>		 
</table> 
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->