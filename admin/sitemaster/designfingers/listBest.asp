<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'##############################################
' History: 2008.03.18 생성
' Description: 디자인 핑거스 최근 3개월 간 베스트 코멘트
'##############################################
 Dim clsDF,clsDFCode
 Dim arrList, intLoop
 Dim arrCode
  
 	
'//리스트 가져오기	
 set clsDF = new CDesignFingers
 	arrList = clsDF.fnGetBestComment 	
 set clsDF = nothing
 
 '//핑거스구분(10)에 해당하는 코드내용 배열에 넣기
 set clsDFCode = new CDesignFingersCode
 	arrCode = clsDFCode.fnGetCommCode(10)	
 set clsDFCode =nothing

 	
%>
<script language="javascript">
<!--
	function jsSearch(){
		document.frmSearch.submit();
	}
	
		
	function jsPopCode(){
		var winCode;
		winCode = window.open('popManageCode.asp','popCode','width=400,height=600');
		winCode.focus();
	}
	
 	function jsSetFile(iDFS){   
 	 var winfile = window.open('','setfile','width=1,height=1');	
 	 	 document.frmFile.iDFS.value = iDFS;
		 document.frmFile.target 	= "setfile";
		 document.frmFile.action 	= "<%=uploadUrl%>/chtml/make_designfingers_FlashText.asp";
		 document.frmFile.submit(); 
		
	 winfile.focus();			 
	}
	
	 //이미지첨부
 function jsPopAddImg(sFolder,sImgID){
 document.domain ="10x10.co.kr";	
 	var chkIcon = 0;
 	var winImg;
 	var sImgURL;
 	 	 	
 		sImgURL = eval("document.frmBest.img"+sFolder+sImgID).value; 	 	
 		winImg = window.open('popAddImage.asp?sF='+sFolder+'&sID='+sImgID+'&chkI='+chkIcon+'&sIU='+sImgURL,'popImg','width=380,height=150');
 		winImg.focus();
 }
 
 	//베스트 선정
 	function jsSetBest(){
 		var frm = document.frmBest;
 		var arrDFS = "";
 		
 		if(typeof(frm.chkID) == "undefined"){
 			alert("선택된 ID가 없습니다.");
 			return;
 		}
 		 		
 		if(typeof(frm.chkID.length) == "undefined"){ 		
 			if(frm.chkID.checked){  				
	 			arrDFS = frm.chkID.value;	 	
	 		}
	 	}else{			 	
	 		for(i=0;i<frm.chkID.length;i++){
	 			if(frm.chkID[i].checked){ 
	 				if(arrDFS ==""){
	 					arrDFS = frm.chkID[i].value;
	 				}else{
	 					arrDFS = arrDFS +"," +frm.chkID[i].value;
	 				}
	 			}
	 		}	 		
	 			
	 	}	
 		
 		if(arrDFS==""){
 		 alert("ID를 선택해 주세요");	
 		 return;
 		}
 		
 		 var winfile = window.open('','setfile','width=1,height=1');	 	 	
		 document.frmBest.target 	= "setfile";
		 document.frmBest.action 	= "<%=uploadUrl%>/chtml/make_designfingers_BestJS.asp?menupos=<%=menupos%>&arrDFS="+arrDFS;
		 document.frmBest.submit(); 
		
	 	winfile.focus();	
 	}
//-->
</script>
 
<table width="800" border="0" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td colspan="2"> + 현재 베스트 리스트<p>
		<script language="JavaScript" src="<%=staticImgUrl%>/chtml/js/designfingers_Best.js"></script>
	</td>	
</tr>
<tr>
	<td colspan="2"><hr width="100%"></td>
</tR>
<tr>
	<td>+ 최근 3개월 리스트, 총코멘트수 순 정리
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
	    <tr height="40" valign="bottom">       
	        <td align="left">
				<input type="button" value="선택ID 베스트선정" class="button" onClick="jsSetBest();">
			</td>
			<td align="right">	
				<input type="button" class="button" value="핑거스리스트" onClick="location.href='listDF.asp?menupos=<%= menupos %>'">
				<% if C_ADMIN_AUTH then %><input type="button" class="button" value="코드관리" onclick="javascript:jsPopCode();"><%END IF%>				
			</td>
		</tr>			
		</table>
	</td>
</tr>
<tr>
	<td colspan="2"> 
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a"  bgcolor="#CCCCCC">	
		<form name="frmBest" method="post" action="">			
		<tr bgcolor="#EFEFEF">
			<td width="40" align="center" nowrap>선택</td>
			<td width="40" align="center" nowrap>ID	</td>
			<td width="60" align="center" nowrap>구분</td>			
			<td align="center">제목</td>
			<td width="60" align="center" nowrap>당첨발표일</td>
			<td width="60" align="center" nowrap>등록일</td>
			<td width="60" align="center" nowrap>총코멘트수</td>
			<td width="150" align="center" nowrap>배너</td>
		</tr>
		<%IF isArray(arrList) THEN%>
		<% For intLoop =0 To UBound(arrList,2) %>	
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="chkID" value="<%=arrList(0,intLoop)%>"></td>
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=fnGetCodeArrDesc(arrCode,arrList(1,intLoop))%></td>			
			<td align="left" ><a href="regDF.asp?iDFS=<%=arrList(0,intLoop)%>&menupos=<%= menupos %>"><%=arrList(2,intLoop)%></a></td>
			<td align="center" ><%=arrList(3,intLoop)%></td>
			<td align="center"><%=FormatDate(arrList(5,intLoop),"0000.00.00")%></td>
			<td align="center" ><%=arrList(7,intLoop)%></td>
			<td align="center"><%IF arrList(6,intLoop) <> "" THEN%><img src="<%=arrList(6,intLoop)%>" width="150"><%END IF%>
			<input type="button" value="등록" class="button" onClick="javascript:jsPopAddImg('banner',<%=arrList(0,intLoop)%>);">
			<input type="hidden" name="imgbanner<%=arrList(0,intLoop)%>" value="<%=arrList(6,intLoop)%>">
			</td>
		</tr> 
		<% Next%>
		<%ELSE%>
		<tr bgcolor="#FFFFFF">
			<td colspan="8" align="center">등록된 내역이 없습니다.</td>
		</tr>
		<%END IF%>	
		</form>
		</table>
	</td>		
</tr>
</table>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->