<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 계정과목 내용  리스트
' History : 2011.03.09 정윤정  생성
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp" -->
<%
Dim clsAcc 
Dim arrList, intLoop
Dim ipcateidx, icateidx,sACCUSECD,sACCNM,sisNoSet 
Dim iCurrPage, iPageSize,iTotCnt,iTotalPage
Dim iCpcateidx,iCcateidx

ipcateidx = requestCheckvar(Request("selCL"),10)
IF ipcateidx = "" THEN ipcateidx = 0
icateidx 	= requestCheckvar(Request("iCS"),10)
IF icateidx = "" THEN icateidx = 0
sACCUSECD = requestCheckvar(Request("sAUCD"),15)
sACCNM 		= requestCheckvar(Request("sANM"),50)
sisNoSet 	= requestCheckvar(Request("chkNS"),1)
iCurrPage = requestCheckvar(Request("iCP"),10)
IF iCurrPage = "" THEN iCurrPage = 1
iPageSize = 30

Set clsAcc = new CAccCategory
	clsAcc.FACCPCateIdx =  ipcateidx
	clsAcc.FACCCateIdx  =  icateidx 
	clsAcc.FACCUSECD    =  sACCUSECD   
	clsAcc.FACCNM       =  sACCNM      
	clsAcc.FisNoSet     =  sisNoSet    
	clsAcc.FCurrPage     =  iCurrPage    
	clsAcc.FPageSize     =  iPageSize    
	arrList = clsAcc.fnGetACCCDList
	iTotCnt	= clsAcc.FTotCnt
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>  
<script  type="text/javascript">
<!--
// 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	   
//과목 관리
function jsMngCategory(){
	var winC = window.open("categoryList.asp","popC","width=600, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
} 
 
//계정과목별 안분기준관리
function jsSetDivide(acccd){
	var winC = window.open("accDivide.asp?Acccd="+acccd,"popC","width=600, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//검색
function jsSearch(){
		document.frm.iCS.value = $("#selC").val();  
		document.frm.submit();
}
//과목 변경 =========================================================================================================
$(document).ready(function(){
	$("#selCL").change(function(){
		var iValue = $("#selCL").val(); 
		var url="/admin/approval/accCategory/ajaxCate.asp";
		 var params = "sVar=selC&selCL="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divC").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
 
	$("#selCCL").change(function(){
		var iValue = $("#selCCL").val(); 
		var url="/admin/approval/accCategory/ajaxCate.asp";
		 var params = "sVar=selCC&selCL="+iValue;  
		   
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divCC").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});
//과목 선택
function jsSetCategory(){
	if($("#selCCL").val() ==0){
		alert("대과목을 선택해주세요");
		return;
	}
	if($("#selCC").val() ==0){
		alert("중과목을 선택해주세요");
		return;
	}
	
	var ischecked =false;
    
    for (var i=0;i<frmReg.elements.length;i++){
		//check optioon
		var e = frmReg.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}
	
	if (!ischecked){
	    alert('선택 내역이 없습니다.');
	    return;
	}
	
	if (confirm('등록하시겠습니까?')){  
			frmReg.iccidx.value = $("#selCC").val();
 	    frmReg.submit();
 	}
}
 
  
	
function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

	function checkThis(comp){
    AnCheckClick(comp)
} 

//카테고리삭제
function jsDelCate(iValue){
	if(confirm("선택하신 계정의 과목을 삭제하시겠습니까?")){
		document.frmDel.hidCDIdx.value = iValue;
		document.frmDel.submit();
	}
}
//-->
</script>
<form name="frmDel" method="post" action="procCategory.asp">
	<input type="hidden" name="hidM" value="D">
	<input type="hidden" name="hidCDIdx" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="iCS" value="<%=icateidx%>">
	<input type="hidden" name="selCL" value="<%=ipcateidx%>">
</form>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td><form name="frm" method="get" action="index.asp" style="margin:0px;">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iCS" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					 과목:
				 	<select name="selCL"  id="selCL">
					<option value="0">--선택--</option>
					<%clsAcc.sbGetOptAccCategory 1,0,ipcateidx %>			
					</select> 
					>
					<span id="divC">
					<select name="selC"  id="selC">
					<option value="0">--선택--</option>
					<% IF ipcateidx > 0 THEN
					clsAcc.sbGetOptAccCategory 2,ipcateidx,icateidx 
						END IF
					%>			
					</select> 
					</span> 
					&nbsp;&nbsp;
					계정과목명: <input type="text" name="sANm" value="<%=sACCNm%>" size="20">
					&nbsp;&nbsp;
					계정과목번호: <input type="text" name="sAUCD" value="<%=sACCUseCD%>" size="15">
					&nbsp;&nbsp;
					<input type="checkbox" name="chkNS" value="Y" <%IF sIsNoSet ="Y" THEN%>checked<%END IF%>> 과목 미지정 계정만
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr> 
		</table>
	</form>
	</td>
</tr>  
<tr>
	<td><hr width="100%"></td>
</tr>
<form name="frmReg" method="post" action="procCategory.asp" style="margin:0px;">
	<input type="hidden" name="hidM" value="S">
	<input type="hidden" name="iccidx" value=""> 
	<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td>	<input type="button" class="button" value="과목 관리" onClick="jsMngCategory();">   
		 &nbsp;|&nbsp; 
	 과목: 
				 	<select name="selCCL"  id="selCCL">
					<option value="0">--선택--</option>
					<%clsAcc.sbGetOptAccCategory 1,0,iCpcateidx %>			
					</select> 
					>
					<span id="divCC">
					<select name="selCC"  id="selCC">
					<option value="0">--선택--</option>
					<% IF iCpcateidx > 0 THEN
					clsAcc.sbGetOptAccCategory 2,iCpcateidx,iCcateidx 
						END IF
					%>			
					</select> 
					</span> 
					<input type="button" class="button" value="과목 등록" onClick="jsSetCategory();"> : 선택  과목에 선택계정 등록
	</td>
</tr>
<%Set clsAcc = nothing%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td> 
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					검색결과 : <b><%=iTotCnt%></b> &nbsp;
					페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>선택</td>
				<td>과목</td>
				<td>계정과목코드</td> 
				<td>계정과목명</td> 
				<td>매출처</td> 
				<td>안분기준</td> 
				<td>처리</td> 
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" <%IF not isNull(arrList(3,intLoop)) THEN %>disabled<%END IF%>></td>
				<td><%IF not isNull(arrList(3,intLoop)) THEN %><%=arrList(6,intLoop)%> > <%=arrList(4,intLoop)%><%END IF%></td>			
				<td><%=arrList(1,intLoop)%></td>	
				<td><%=arrList(2,intLoop)%></td>	
				<td><%if arrList(8,intLoop) then %>10x10<%end if%>&nbsp;
					<%if arrList(9,intLoop) then %>제휴<%end if%>
				</td> 
				<td><%=arrList(10,intLoop)%></td>
				<td>
				
				<input type="button" value="안분기준관리" class="button" onClick="jsSetDivide(<%=arrList(0,intLoop)%>);"
				<%IF   isNull(arrList(3,intLoop)) THEN %>disabled<%END IF%>
				>
				&nbsp;
				
				<input type="button" value="삭제" class="button" onClick="jsDelCate(<%=arrList(7,intLoop)%>);">
				</td>	 
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="4">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
		</table>	 
	</td> 
</tr> 
</form>	
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>        
					    </tr>        
					</table>
				</td>
			</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
 



	