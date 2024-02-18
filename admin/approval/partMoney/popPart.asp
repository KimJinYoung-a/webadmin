<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : manager regist
' History : 2011.03.26 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp" -->
<%
Dim clsPart,istep1partidx , istep2partidx,ieapppartidx
Dim ieappDepth,seappPartName,ipartSort,blnUsing ,arrData2
Dim sMode,menupos
Dim arrData, intD
  
ieapppartidx= requestCheckvar(Request("iepidx"),10) 
menupos		= requestCheckvar(Request("menupos"),10) 

sMode = "I"

Set clsPart= new CpartMoneyCls
	clsPart.Fstep1partidx = 0
	clsPart.Fstep2partidx = 0
	clsPart.FeappDepth = 1
	arrData = clsPart.fnGetPartList 	
IF ieapppartidx <> "" THEN
	sMode ="U"
	clsPart.Feapppartidx = ieapppartidx
	clsPart.fnGetPartData	
	  
	ieappDepth		=clsPart.FeappDepth	
	istep1partidx 	=clsPart.Fstep1partidx  
	istep2partidx 	=clsPart.Fstep2partidx  
	seappPartName  	=clsPart.FeappPartName   
	ipartSort		=clsPart.FpartSort
	blnUsing       	=clsPart.FisUsing 
	 
	IF istep1partidx > 0 and ieappDepth = 3 THEN 
		clsPart.Fstep1partidx = istep1partidx 
		clsPart.Fstep2partidx = 0
		clsPart.FeappDepth	  = 2	
		arrData2 = clsPart.fnGetPartList  
	END IF	    
 END IF     
 Set clsPart= nothing 
 %>

<script type="text/javascript" src="/js/ajax.js"></script>	
<script language="javascript">
<!-- 
   initializeReturnFunction("processAjax()");
   initializeErrorFunction("onErrorAjax()"); 
    
    function processAjax(){
        var reTxt = xmlHttp.responseText;  
        document.all.sp2.innerHTML = reTxt; 
    }
    
    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }
    
    //선택한 부서에 대한 하위 부서리스트 가져오기 Ajax
    function jsSetPart(){ 
      var is1  = document.frm.selp1.value;  
        initializeURL('ajaxPart.asp?is1='+is1+'&hidM=R&idp=2');
    	startRequest(); 
    }
	
	//등록전 필드 체크
	function jsSubmit(){
	 if(document.frm.sPN.value==""){
	 alert("부서명을 입력해주세요");
	 document.frm.sPN.focus();
	 return false;
	 }
	  
	 return true;
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>자금관리 부서등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frm" method="post" action="procPart.asp" OnSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="iepidx" value="<%=ieapppartidx%>"> 
		<input type="hidden" name="menupos" value="<%=menupos%>">		
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">부서 IDX</td>
			<td bgcolor="#FFFFFF" width="380"><%=ieapppartidx%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">위치지정<br>(상위부서)</td>
			<td bgcolor="#FFFFFF" width="380">  
			<span id="sp1">
			<select name="selp1" onChange="jsSetPart()"> 
			<option value="0">--최상위--</option> 
			<%IF isArray(arrData) THEN
				For intD = 0 To UBound(arrData,2)
			%>
				<option value="<%=arrData(0,intD)%>" <%IF Cstr(istep1partidx) = Cstr(arrData(0,intD)) THEN%>selected<%END IF%>><%=arrData(4,intD)%></option>
			<%	
				Next
			  END IF
			%> 
			</select> 
			</span>
			<span id="sp2">
			<%IF isArray(arrData2) THEN %> 
				<select name="selp2">  
				<option value="0">----</option>
			<%
			For intD = 0 To UBound(arrData2,2)
			%>
				<option value="<%=arrData2(0,intD)%>" <%IF Cstr(istep2partidx) = Cstr(arrData2(0,intD)) THEN%>selected<%END IF%>><%=arrData2(4,intD)%></option>
			<%	
				Next
			%>
				</select>
			<%	
			  END IF	
			 %>
			</span> 
			</td>
		</tr> 		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">부서명</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sPN" value="<%=seappPartName%>" size="20"></td>
		</tr> 
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">정렬순서</td>
			<td bgcolor="#FFFFFF"><input type="text" name="iPS" value="<%=ipartSort%>" size="4"></td>
		</tr> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">사용유무</td>
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoU" value="1" checked>사용 <input type="radio" name="rdoU" value="0">사용안함</td>
		</tr>	
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="등록" class="button"></td>
</tr>
</form>
</table>
</body>
</html> 
<!-- #include virtual="/lib/db/dbclose.asp" --> 
	