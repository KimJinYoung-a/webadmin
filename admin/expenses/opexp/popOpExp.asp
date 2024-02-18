<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : 운영비관리  내용
' History : 2011.05.30 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<%
Dim sMode
Dim clsPart, arrType , clsAccount, arrAccount 
Dim iOpExpPartIdx, iPartTypeIdx, sOpExpPartName, blnUsing,arrPartsn, intLoop, iPartsn
Dim sPartTypeName
Dim intY, dYear, intM, dMonth
iOpExpPartIdx = requestCheckvar(Request("hidOEP"),10) 
sMode ="I"

  '구분값 가져오기
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList 
Set clsPart = nothing

set clsAccount = new COpExpAccount
	arrAccount = clsAccount.fnGetAccountAll
set clsAccount = nothing  
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<script language="javascript">	
<!--
 	//등록
 	function jsPartSubmit(){
 		if(document.frmReg.selPT.value==0 && document.frmReg.sPTN.value==""){
 		alert("구분명을 등록해주세요");
 		return;
 		}
 		
 		if( document.frmReg.sPN.value==""){
 		alert("운영비 관리팀명을 입력해주세요");
 		return;
 		}
 		
 		document.frmReg.submit();
 	}
	  
	  //구분 
	  function jsChPT(iValue){
	  if (iValue==0){
	  	document.all.divPT.style.display = "";
	  	}else{
	  	document.all.divPT.style.display = "none";
	  	}
	  }
	  
	  //부서 추가
	  function jsAddPart(){
	    var winPart = window.open("popAddPart.asp","popPart","width=600, height=600");
	    winPart.focus();
	  }
	  
	  //선택부서 삭제
	  function jsDelPart(iValue){
	    var arrValue = document.frmReg.hidPsn.value.split(",");  
	    if(typeof(arrValue.length)=="undefined"){
	    	document.frmReg.hidPsn.value  = ""
	    }else{
	    	if(arrValue[0] == iValue){
	    		document.frmReg.hidPsn.value  = document.frmReg.hidPsn.value.replace(iValue,"");	
	    	}else{
	    	 	document.frmReg.hidPsn.value  = document.frmReg.hidPsn.value.replace(","+iValue,"");
	    	}
	    } 
	  	eval("document.all.dP"+iValue).outerHTML = "";
	  }
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>운영비 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popAccount.asp"> 
			<input type="hidden" name="iCP" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">
					 날짜 :
					 <select name="selY">
					 <%For intY = Year(date()) To 2011 STEP -1%>
					 <option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					 </select>년
					  <select name="selM">
					 <%For intM = 1 To 12%>
					 <option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
					 </select>월
					 &nbsp;&nbsp;&nbsp;
					 운영비관리팀:
					 <select name="selPT">
					 <option value="">--선택--</option>
					 <% sbOptPartType arrType,ipartTypeIdx%>
					 </select>
					 <select name="selP">
					 </select> 
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
		</form>
		</table>
	</td>
</tr> 
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frmReg" method="post" action="procPart.asp"> 
		<input type="hidden" name="hidM" value="<%=sMode%>">  
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
		 	<td>날짜</td>  
			<td>계정</td>  
			<td>업체명</td>  
			<td>입금</td>  
			<td>출금</td>  
			<td>적요(상세내역)</td>   
			<td>처리</td>  	  
		</tr> 
		<tr bgcolor="#FFFFFF"  align="center">
		 	<td><input type="text" name="iD" size="2"></td>  
			<td><select name="selA">
				<% sbOptAccount arrAccount, ""%>
				</select></td>  
			<td><input type="text" name="sO" size="20"></td>  
			<td><input type="text" name="mIn" size="10"></td>  
			<td><input type="text" name="mOut" size="10"></td> 
			<td><input type="text" name="sDC" size="40" maxlength="200"></td> 
			<td><input type="button" class="button" value="추가"></td>  	   	  
		</tr> 
		</form>
		</table>	
	</td> 
</tr>  
</table>
<!-- 페이지 끝 -->
</body>
</html>
 



	