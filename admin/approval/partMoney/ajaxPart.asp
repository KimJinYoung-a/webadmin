<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim istep1partidx,istep2partidx, idepth
Dim clsPart,arrList,intLoop
Dim sMode

sMode =requestCheckvar(Request("hidM"),1)
idepth = requestCheckvar(Request("iDP"),1)
istep1partidx = requestCheckvar(Request("is1"),10)
IF istep1partidx = "" THEN istep1partidx = 0
	
istep2partidx = requestCheckvar(Request("is2"),10)
IF istep2partidx = "" THEN istep2partidx = 0 
 
Set clsPart = new CpartMoneyCls
	clsPart.Fstep1partidx = istep1partidx
	clsPart.Fstep2partidx = istep2partidx 
	clsPart.FeappDepth	  = idepth	
	arrList = clsPart.fnGetPartList  
Set clsPart = nothing 
 
IF sMode = "R" THEN '자금관리 부서 등록시 사용
%>
	<%IF isArray(arrList) THEN%>
	&nbsp;>&nbsp;<select name="selp2" id="selp2">  
	<option value="0">----</option>
	<% For intLoop = 0 To UBound(arrList,2)%>
	<option value="<%=arrList(0,intLoop)%>"><%=arrList(4,intLoop)%></option>
	<% 	Next %>  
	</select> 
	<% END IF %> 
<%ELSEIF sMode = "S" THEN '전자결재에서 부서 선택시 사용
%> 
<script language="javascript">
<!-- 
$(document).ready(function(){
	$("#selp2").change(function(){  
		var selp1 = $("#selp1").val();
		 var selp2 = $("#selp2").val();
		 var url="/admin/approval/partMoney/ajaxPart.asp";
		 var params = "hidM=S&iDP=3&is1="+selp1+"&is2="+selp2;
		 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){  
		 		$("#sp3").html(args);	 
		 	}, 
		 	error:function(e){ 
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
}); 
//-->
</script> 
	<select name="selp<%=idepth%>" id="selp<%=idepth%>">  
	<option value="0">--선택--</option> 
	<%IF isArray(arrList) THEN%> 
	<% For intLoop = 0 To UBound(arrList,2)%>
	<option value="<%=arrList(0,intLoop)%>"><%=arrList(4,intLoop)%></option>
	<% 	Next %> 
	<% END IF %>  
	</select>  
<%ELSE '리스트 view%> 
<select name="selp2" id="selp2">  
<option value="0">--선택--</option>
	<%IF isArray(arrList) THEN%> 
	<% For intLoop = 0 To UBound(arrList,2)%>
	<option value="<%=arrList(0,intLoop)%>"><%=arrList(4,intLoop)%></option>
	<% 	Next %>  
	<% END IF %>
	</select>  
<%END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->