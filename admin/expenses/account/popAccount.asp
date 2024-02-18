<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 계정과목 리스트 - 운영비 선택
' History : 2011.05.30 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExparapCls.asp"-->
<%
Dim clsAccount, arrList, intLoop  	
Dim sarap_nm
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		 
 	sarap_nm =  requestCheckvar(Request("sAN"),50)
 	
Set clsAccount = new COpExpAccount
	clsAccount.Farap_nm 	= sarap_nm 
	clsAccount.FCurrPage 	= iCurrPage
	clsAccount.FPageSize 	= iPageSize
	arrList = clsAccount.fnGetAccountList 	
	iTotCnt = clsAccount.FTotCnt
Set clsAccount = nothing
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>  
<script language="javascript">
<!--
	 
	 function jsAddOE(){
	 var chkValue="";
	 if(typeof(document.frmReg.chkcd)=="undefined"){  
	 	return;
	 }
	 
	 if(typeof(document.frmReg.chkcd.length)=="undefined"){  
	  	if(document.frmReg.chkcd.checked){ 
	  	 chkValue = document.frmReg.chkcd.value; 
	  	}
	  } 
	 
	 for(i=0;i<document.frmReg.chkcd.length;i++){
	  if(document.frmReg.chkcd[i].checked){
	   if(chkValue==""){
	   	chkValue = document.frmReg.chkcd[i].value;
	   }else{
	  	chkValue = chkValue +","+document.frmReg.chkcd[i].value;
	  	}
	  	}
	 }
	 
	 if(chkValue==""){
	 	alert("추가하실 수지항목을 선택해주세요");
		 return;
	 }
	
	 document.frmReg.hidccd.value = chkValue;
	 document.frmReg.submit();
	 
	 }
	 
	 // 페이지 이동
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>수지항목선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popAccount.asp"> 
			<input type="hidden" name="iCP" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">
					수지항목: 
					 <input type="text" name="sAN" size="20" maxlenght="30" value="<%=sarap_nm%>">
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
		</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<tr>
	<td align="right"><input type="button" class="button" value="수지항목 추가" onClick="jsAddOE();"></td>
</tr>
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
			<form name="frmReg" method="post" action="procAccount.asp"> 
			<input type="hidden" name="hidM" value="I">
			<input type="hidden" name="hidccd" value="">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			 	<td>선택</td>  
			 	<td>운영비코드</td>  
				<td>수지항목</td>  
				<td>연결계정과목</td> 
				<td>erpcode</td>   
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 bgcolor="<%IF not isNull(arrList(5,intLoop)) THEN%><%=adminColor("green")%><%ELSE%>#FFFFFF<%END IF%>">	
				<td  align="center"><input type="checkbox" value="<%=arrList(0,intLoop)%>" name="chkcd" <%IF not isNull(arrList(5,intLoop)) THEN%>disabled<%END IF%>></td>
				<td  align="center"><%=arrList(5,intLoop)%></td>
				<td>[<%=arrList(0,intLoop)%>] <%=arrList(1,intLoop)%></td>
				<td>[<%=arrList(3,intLoop)%>] <%=arrList(4,intLoop)%></td>			
				<td><%=arrList(2,intLoop)%></td>	 
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="5">등록된 내용이 없습니다.</td>	
			</tr>
			<%END IF%>
			</form>
		</table>	
	</td> 
</tr> 	
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
 



	