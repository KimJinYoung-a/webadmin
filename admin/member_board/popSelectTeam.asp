<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  공지사항 팀 등록
' History : 2016.03.30 정윤정 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim did, iLoop, imax , chkdid
did = split(requestCheckvar(Request("did"),200),",")
imax = ubound(did) 
iLoop = 0
chkdid = 0
%>
<script type="text/javascript">
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
	
	//선택팀 등록처리
	function jsSetTeam(){
	    var strTeam;
	    strTeam = "";
	    if(document.all.cid.length!=undefined){
	        for(i=0;i<document.all.cid.length;i++){ 
	            if(document.all.cid[i].checked){
	               strTeam = strTeam + "<div id='dvSTeam"+i+"'><label><input type='hidden' name='did' value='"+document.all.cid[i].value+"'>" 
	               strTeam = strTeam + document.all.cnm[i].value+" <a href='javascript:jsTeamDel("+i+");'>[x]</a></label></div>"
	                 
	            }
	        }
	    }
	  
	    opener.document.all.dvTeam.innerHTML=strTeam;
	    self.close();
	}
</script>
<div id="divUL" style="padding:5px;border:gray solid 1px;width:300px;height:610px;overflow-y:auto;"> 
<%
dim cTenDepartment, i
dim cid1,cid2,cid3,cid4,oldcid1,oldcid2,oldcid3,oldcid4,imgJoin, ncid1, ncid2, ncid3, ncid4
dim department_id
set cTenDepartment = new CTenByTenDepartment

cTenDepartment.FPageSize = 500
cTenDepartment.FCurrPage = 1
cTenDepartment.FRectUseYN = "Y"

cTenDepartment.GetList
for i = 0 to cTenDepartment.FResultcount - 1
    if imax>=0  then
        chkdid = trim(did(iLoop))
     end if   
    department_id = cTenDepartment.FItemList(i).Fcid
     
    cid1 = cTenDepartment.FItemList(i).Fcid1
	cid2 = cTenDepartment.FItemList(i).Fcid2
	cid3 = cTenDepartment.FItemList(i).Fcid3
	cid4 = cTenDepartment.FItemList(i).Fcid4

    if isnull(cid2) then cid2 = 0 
	if isnull(cid3) then cid3 = 0 
	if isnull(cid4) then cid4 = 0 	
	
	'//부서가 틀려지면 라인이미지가 틀려진다
	if i < cTenDepartment.FResultcount - 1  THEN 
	    ncid1 = cTenDepartment.FItemList(i+1).Fcid1
	    ncid2 = cTenDepartment.FItemList(i+1).Fcid2
	    ncid3 = cTenDepartment.FItemList(i+1).Fcid3
	    ncid4 = cTenDepartment.FItemList(i+1).Fcid4
	    
	    if isnull(ncid2) then ncid2 = 0 
	    if isnull(ncid3) then ncid3 = 0 
	    if isnull(ncid4) then ncid4 = 0 
	    
		IF  (cid2 <> ncid2 and cid3 <>  ncid3 and cid4=0)  or (cid4<>0 and cid3 <>  ncid3)  THEN
			imgJoin = "joinbottom.gif"
		ELSE	
			imgJoin = "join.gif"
		END IF
	else
	     ncid1 = 0
         ncid2 = 0
         ncid3 = 0
         ncid4 = 0
    	imgJoin = "joinbottom.gif"
	end if	  
	
	if cid1 <> oldcid1 then	
	%>		
	<% if i <> 0 then%> 
		<% if oldcid2 <> 0 then%>
			<% if oldcid3 <> 0 then%>  
			</div>
			<%end if%>
		</div>
		<%end if%>
	</div>
	<% end if%> 
	 
	<div id="divP1<%=cid1%>" ><a href="javascript:jsOpenClose('1','<%=cid1%>');" style="text-decoration:none;"><img src="/images/Tminus.png" align="absmiddle" id="Timg<%=cid1%>" style="border:0px"> 
	        <img src="/images/dtree/openfolder.png" align="absmiddle" id="Fimg<%=cid1%>" style="border:0px"> 
	    <%=cTenDepartment.FItemList(i).FdepartmentName1%></a>
    </div>  	 
	<div id="divB1<%=cid1%>" style="display:;cursor:hand;">	 
<%  end if%>		
<% if cid2  <>  oldcid2 and cid2 <> 0 then	 
	 if oldcid2<> 0 then
%>	 
	<% if oldcid3 <> 0 then%>  
	</div>
	<%end if%>
	</div>  
<%	end if%>		 
	    <div id="divP2<%=cid2%>" style="padding:0 0 0 1;" ><img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle" style="border:0px">
			<%if cid2= ncid2 then%>
			<a href="javascript:jsOpenClose('2','<%=cid2%>');" style="text-decoration:none;">
			<img src="/images/Tminus.png" align="absmiddle" id="Timg<%=cid2%>" style="border:0px">
			<img src="/images/dtree/openfolder.png" align="absmiddle" id="Fimg<%=cid2%>" style="border:0px">
			<%=cTenDepartment.FItemList(i).FdepartmentName2%></a>
			<%else%>
			 <img src="/images/dtree/<%=imgJoin%>" align="absmiddle"><input type="checkbox" name="cid" value="<%=department_id%>" <%if cstr(chkdid) =cstr(department_id) then %>checked<%end if%>>
			 <input type="hidden" name="cnm" value="<%=cTenDepartment.FItemList(i).FdepartmentNameFull%>">
			 <%=cTenDepartment.FItemList(i).FdepartmentName2%> 
			<%end if%> 
	    </div>  	 
		<div id="divB2<%=cid2%>" style="display:;cursor:hand;">
 <%end if%>	
  <% if cid3  <>  oldcid3 and cid3 <> 0 then	 
		 if oldcid3<> 0 then
	%>	 
		 
		</div>  
<%  end if%>		 
		<div id="divP3<%=cid3%>" style="padding:0 0 0 1;" >
			<img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
			<img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle"> 
			<%if cid3= ncid3 then%>
			<a href="javascript:jsOpenClose('3','<%=cid3%>');" style="text-decoration:none;">
			    <img src="/images/Tminus.png" align="absmiddle" id="Timg<%=cid3%>" style="border:0px"><img src="/images/dtree/openfolder.png" align="absmiddle" id="Fimg<%=cid3%>" style="border:0px">
			<%=cTenDepartment.FItemList(i).FdepartmentName3%> </a>
			<%else%>
			    <img src="/images/dtree/<%=imgJoin%>" align="absmiddle"><input type="checkbox" name="cid" value="<%=department_id%>" <%if cstr(chkdid) =cstr(department_id) then %>checked<%end if%>>
			    <input type="hidden" name="cnm" value="<%=cTenDepartment.FItemList(i).FdepartmentNameFull%>">
			    <%=cTenDepartment.FItemList(i).FdepartmentName3%> 
			<%end if%> 
	    </div>  	 
		<div id="divB3<%=cid3%>" style="display:;cursor:hand;">		
	<%end if%>	 
	<% if cid4  <>  oldcid4 and cid4 <> 0 then	  %>
	 	 
		<div id="divP4<%=cid4%>" style="cursor:hand;padding:0 0 0 1;">
			<img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
			<img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
			<img src="<%IF department_id<>cTenDepartment.FItemList(cTenDepartment.FResultcount - 1).Fcid then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle">
			<img src="/images/dtree/<%=imgJoin%>" align="absmiddle"><input type="checkbox" name="cid" value="<%=department_id%>" <%if cstr(chkdid) =cstr(department_id) then %>checked<%end if%>><%=cTenDepartment.FItemList(i).FdepartmentName4%> 
			<input type="hidden" name="cnm" value="<%=cTenDepartment.FItemList(i).FdepartmentNameFull%>">
	    </div>  	 
				 
<%end if%>	   
   	<% 
   	    if cstr(chkdid) = cstr(department_id) then 
   	        if imax > iLoop then
   	        iLoop = iLoop + 1 
   	        end if
   	    end if
   	    oldcid1 = cid1
		oldcid2 = cid2
		oldcid3 = cid3
		oldcid4 = cid4
       next
   	%> 
   	</div>
  </div> 	
   	<%set cTenDepartment = nothing%>
 <div style="padding-top:10px;text-align:right;"><input type="button" class="button" value="등록" style="width:100px;" onClick="jsSetTeam();"> </div>