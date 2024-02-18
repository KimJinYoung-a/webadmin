<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
	Dim arrList,intLoop,oMember
	Dim sUsername,sWorkerID ,sType
  Dim imgJoin,part_sn, oldpart_sn
	dim department_id,  departmentname,user_cid1, user_cid2, user_cid3, user_cid4
	 
	sType 		=  requestCheckvar(Request("sType"),3)
	sWorkerID=  requestCheckvar(Request("workerid"),32)
	department_id = requestCheckVar(Request("department_id"),10) 
	sUsername =  requestCheckvar(Request("sUN"),32)
   
Set oMember = new CTenByTenMember
	oMember.Fusername 		= sUsername
	arrList = oMember.fnGetUserTreeListNew
		
	if 	sWorkerID = "" then
	oMember.Fdepartment_id =  department_id 
	else
	oMember.Fuserid 			 =  sWorkerID
	end if
	oMember.fnGetDepartmentInfoPID
	department_id		= oMember.Fdepartment_id
 	departmentname = oMember.FdepartmentNameFull
 	user_cid1						= oMember.Fcid1
 	user_cid2						= oMember.Fcid2
 	user_cid3						= oMember.Fcid3
 	user_cid4						= oMember.Fcid4 
   
set oMember = nothing
 
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"> </script> 
<script type="text/javascript">
window.document.domain = "10x10.co.kr";
	//트리모양 클릭에 따라 변경
	function jsOpenClose(cValue,iValue){  
		if(eval("document.all.divB"+cValue+iValue).style.display=="none"){
			eval("document.all.Fimg"+iValue).src = "/images/dtree/openfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tminus.png";
			eval("document.all.divB"+cValue+iValue).style.display=""; 
		}else{
			eval("document.all.Fimg"+iValue).src = "/images/dtree/closedfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tplus.png";
			eval("document.all.divB"+cValue+iValue).style.display="none"; 
		}
	}
	
	//등록처리 
	function jsSetId(sVal, sUserID, sUserNm){
		opener.eval("document.frmEvt.s"+sVal+"Id").value = sUserID;
		opener.eval("document.frmEvt.s"+sVal+"Nm").value = sUserNm;
		self.close();
	}
	
	
	//검색 
$(document).ready(function(){
	$("#btnSearch").click(function(){
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value ="";  
		var username = escape($("#sUN").val()); 
		 var url="ajaxUserList.asp";
		 var params = "sUN="+username+"&sType=<%=sType%>"; 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){ 
		 		$("#divUL").html(args);
		 	},

		 	error:function(e){
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 });
	});   
	
	 <%if user_cid1 <> "" and not isnull(user_cid1) then%>  
	 		jsOpenClose(1,<%=user_cid1%>); 
	 <%end if%>		 
	 <%if user_cid2 <> "" and not isnull(user_cid2) then%> 
	   jsOpenClose(2,<%=user_cid2%>); 
	 <%end if%>  
	 <%if user_cid3 <> "" and not isnull(user_cid3) then%> 
	   jsOpenClose(3,<%=user_cid3%>); 
	 <%end if%>  
	 <%if user_cid4 <> "" and not isnull(user_cid4) then%> 
	   jsOpenClose(4,<%=user_cid4%>);   
	 <%end if%>  
}); 

</script>
<form name="frmS" id="frmS" method="post" onsubmit="return false;">
	<input type="hidden" name="hidUI" value="">
	<input type="hidden" name="hidUN" value="">  
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
		<tr>
			<td>사원명: <input type="text" name="sUN" id="sUN" value="<%=sUserName%>"> <input type="button" class="button" value="검색" id="btnSearch"></td>
		</tr>
		<tr>
			<td><%dim cid1, cid2, cid3, cid4, oldcid1, oldcid2, oldcid3, oldcid4 %> 
				<div id="divUL" style="padding:5px;border:gray solid 1px;width:300px;height:500px;overflow-y:auto;"> 
				<%IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2) 
							part_sn = arrList(0,intLoop)
							cid1 = arrList(10,intLoop)
							cid2 = arrList(11,intLoop)
							cid3 = arrList(12,intLoop)
							cid4 = arrList(13,intLoop)
							
							if isnull(cid2) then cid2 = 0 
							if isnull(cid3) then cid3 = 0 
							if isnull(cid4) then cid4 = 0 	
							'//부서가 틀려지면 라인이미지가 틀려진다
							if intLoop < UBound(arrList,2)  THEN 
								IF part_sn <> arrList(0,intLoop+1)  THEN
									imgJoin = "joinbottom.gif"
								ELSE	
									imgJoin = "join.gif"
								END IF
							else
								imgJoin = "joinbottom.gif"
							end if	 
						
						if cid1 <> oldcid1 then	
					%>		
						<% if intloop <> 0 then%> 
							<% if oldcid2 <> 0 then%>
									<% if oldcid3 <> 0 then%> 
										<% if oldcid4 <> 0 then%>
										</div>
										<%end if%>
									</div>
									<%end if%>
							</div>
							<%end if%>
						</div>
						<% end if%> 
						<div id="divP1<%=cid1%>" style="cursor:hand;" onClick="jsOpenClose('1','<%=cid1%>');"><img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid1%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid1%>"> <%=arrList(1,intLoop)%></div>  	 
						<div id="divB1<%=cid1%>" style="display:none;cursor:hand;">	 
					<%end if%>	 
			 	 	<% if cid2  <>  oldcid2 and cid2 <> 0 then	 
			 	 				 if oldcid2<> 0 then
					%>	 
							<% if oldcid3 <> 0 then%> 
								<% if oldcid4 <> 0 then%>
										</div>
										<%end if%>
									</div>
									<%end if%>
								</div>  
					<% 			end if%>		 
							<div id="divP2<%=cid2%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('2','<%=cid2%>');">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid2%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid2%>"> <%=arrList(1,intLoop)%></div>  	 
							<div id="divB2<%=cid2%>" style="display:none;cursor:hand;">		
					<%end if%>	
						<% if cid3  <>  oldcid3 and cid3 <> 0 then	 
			 	 				 if oldcid3<> 0 then
					%>	 
								<% if oldcid4 <> 0 then%>
										</div>
										<%end if%>
								</div>  
					<% 			end if%>		 
							<div id="divP3<%=cid3%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('3','<%=cid3%>');">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid3%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid3%>"> <%=arrList(1,intLoop)%></div>  	 
							<div id="divB3<%=cid3%>" style="display:none;cursor:hand;">		
					<%end if%>	
				<% if cid4  <>  oldcid4 and cid4 <> 0 then	 
			 	 				 if oldcid4<> 0 then
					%>	 
								</div>  
					<% 			end if%>		 
							<div id="divP4<%=cid4%>" style="cursor:hand;padding:0 0 0 1;" onClick="jsOpenClose('4','<%=cid4%>');">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
								<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid4%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid4%>"> <%=arrList(1,intLoop)%></div>  	 
							<div id="divB4<%=cid4%>" style="display:none;cursor:hand;">		
					<%end if%>	 
						<div id="divU" style="padding:0 0 0 1;background:white;" onClick="jsSetId('<%=sType%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>');">
								<%if not isnull(arrList(4,intLoop)) then %>		
							<%if cid4<> 0 then%>
							<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
							<%end if%>
							<%if cid3<> 0 then%>
							<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
							<%end if%>
							<%if cid2<> 0 then%>
							<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
							<%end if%>
							<img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
							<img src="/images/dtree/<%=imgJoin%>" align="absmiddle"> 
							<%=arrList(5,intLoop)%>&nbsp;<%=arrList(8,intLoop)%> <font color="gray">[<%=arrList(4,intLoop)%>]</font> 
							<%end if%>
						</div>
						
				<%		
						oldcid1 = cid1
						oldcid2 = cid2
						oldcid3 = cid3
						oldcid4 = cid4
					Next
				END IF
			%>
				</div>
			</td>
		</tr>
	</table>
</form>