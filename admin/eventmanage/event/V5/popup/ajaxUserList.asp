<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->

<%
dim oMember, arrList,intLoop, sType
dim sUsername,part_sn,oldpart_sn
Dim imgJoin
dim cid1, cid2, cid3, cid4, oldcid1, oldcid2, oldcid3, oldcid4 
 sUsername =  requestCheckvar(Request("sUN"),32) 
 sType=  requestCheckvar(Request("sType"),3)
'사원리스트 가져오기
	Set oMember = new CTenByTenMember
	oMember.Fusername 		= sUsername
	arrList = oMember.fnGetUserTreeListNew
	set oMember = nothing
%>
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
					IF part_sn <> arrList(0,intLoop+1) THEN
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
											<div id="divP1<%=cid1%>" style="cursor:hand;" onClick="jsOpenClose('1','<%=cid1%>');"><img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid1%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid1%>"> <%=arrList(18,intLoop)%></div>  	 
											<div id="divB1<%=cid1%>" style="display:<%if susername = "" then%>none<%end if%>;cursor:hand;">	 
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
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid2%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid2%>"> <%=arrList(19,intLoop)%></div>  	 
												<div id="divB2<%=cid2%>" style="display:<%if susername = "" then%>none<%end if%>;cursor:hand;">		
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
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid3%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid3%>"> <%=arrList(20,intLoop)%></div>  	 
												<div id="divB3<%=cid3%>" style="display:<%if susername = "" then%>none<%end if%>;cursor:hand;">		
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
													<img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=cid4%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=cid4%>"> <%=arrList(21,intLoop)%></div>  	 
												<div id="divB4<%=cid4%>" style="display:<%if susername = "" then%>none<%end if%>;cursor:hand;">		
										<%end if%>	 
										<%if not isnull(arrList(4,intLoop)) then %>							
											<div id="divU" style="padding:0 0 0 1;background:white;" onClick="jsSetId('<%=sType%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>');">
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
											</div>
											<%end if%>
									<%		
											oldcid1 = cid1
											oldcid2 = cid2
											oldcid3 = cid3
											oldcid4 = cid4
										Next
									END IF%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->