<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 문서관리  리스트 - 공통사용
' History : 2011.11.15 정윤정  생성
'	jsSetEdms 스크립트 함수 opener에서 생성해서 선택처리
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<%
Dim clsEdms
Dim arrList, intLoop
Dim icateidx1, icateidx2, sEdmsName
Dim iTotCnt,iPageSize, iTotalPage,iCurrentPage
Dim blnUsing
  blnUsing = 1 '--사용중인 문서만 나오게
	iPageSize = 20
	iCurrentPage = requestCheckvar(Request("page"),10)
	if iCurrentPage="" then iCurrentPage=1
	icateidx1 = requestCheckvar(Request("selC1"),10)
	icateidx2 = requestCheckvar(Request("hidC2"),10) 
 
	if icateidx1 = "" then icateidx1 = 0
	if icateidx2 = "" then icateidx2= 0
	sEdmsName = requestCheckvar(Request("sEN"),60)   

Set clsEdms = new Cedms
	 clsEdms.FCateIdx1	=icateidx1 	
	 clsEdms.FCateIdx2 	=icateidx2 
	 clsEdms.FEdmsName  =sEdmsName 
	 clsEdms.FisUsing 	= blnUsing
	 clsedms.FCurrPage 	= iCurrentPage
	 clsedms.FPageSize 	= iPageSize
	 arrList = clsEdms.fnGetEdmsList 
	 iTotCnt = clsedms.FTotCnt

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script> 
<script language="javascript">
	// 카테고리 ajax =========================================================================================================
	$(document).ready(function(){
	$("#selC1").change(function(){
		var iValue = $("#selC1").val();
		var url="/admin/approval/edms/ajaxCategory.asp";
		 var params = "sMode=CL&ipcidx="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divCL").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});

 function jsSearch(){    
		document.frm.hidC2.value = $("#selC2").val(); 
		document.frm.submit();
	}	
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>문서 선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popGetEdms.asp">   
			<input type="hidden" name="hidC2" value="">
			<tr align="center" bgcolor="#FFFFFF">
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">
						대 카테고리 :
					<select name="selC1">
					<option value="0">전체</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>
					
					중 카테고리 :
					<span id="divCL">
					<select name="selC2" id="selC2">
					<option value="0">전체</option>
				<% 	IF icateidx1 > 0 THEN	'대카테고리 선택 후 중카테고리 선택가능하게
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2 
					END IF
				%>
					</select>
					</span>
					<%Set clsEdms = nothing%>
					</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="jsSearch();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>문서명: <input type="text" name="sEN" value="<%=sEdmsName%>" size="20"></td>
			</tr>				
		</form>
		</table>
	</td>
</tr> 
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
				<td>idx</td>
				<td>문서코드</td>
				<td>대카테고리</td>
				<td>중카테고리</td>  
				<td>문서명</td>
			<td>선택</td>  
		</tr> 
		<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=arrList(0,intLoop)%></td>
		 	<td><%=arrList(7,intLoop)%></td> 
		 	<td><%=arrList(2,intLoop)%></td> 
		 	<td><%=arrList(4,intLoop)%></td> 
		 	<td><%=arrList(6,intLoop)%></td> 
		 	<td><input type="button" class="button" value="선택" onClick="opener.jsSetEdms('<%=arrList(0,intLoop)%>','<%=arrList(6,intLoop)%>');self.close();"> </td>
		</tr>  
	<%	Next
		END IF%>
		</table>	
	</td> 
</tr>  
</table>
<!-- 페이지 시작 -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrentpage mod iPerCnt) = 0 Then
			iEndPage = iCurrentpage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrentpage) then
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
 