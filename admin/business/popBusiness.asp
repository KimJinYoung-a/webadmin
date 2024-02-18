<%@ language=vbscript %>
<% option explicit  %> 
<%
'###########################################################
' Description : 운영비관리  사업자 정보 리스트
' History : 2011.09.26 정윤정  생성
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/Business/BusinessInfoCls.asp"-->
<% 
Dim clsBusi
Dim arrList, intLoop
Dim sBusiNo, sBusiName

 sBusiNo=requestCheckvar(Request("sBNo"),13) 
 sBusiName=requestCheckvar(Request("sBNa"),60) 
 
IF sBusiNo <> "" or sBusiName <> "" THEN
Set clsBusi = new CBsuiness  
	clsBusi.FBusiNo = sBusiNo
	clsBusi.FBusiName = sBusiName
	arrList = clsBusi.fnGetBusinessList 
Set clsBusi = nothing
END IF
%>
<script language="javascript">
<!--
//검색
function jsSearch(){
	document.frmS.submit();
}
 
 //선택
 function jsChoice(iBidx, sBNo){
 	opener.document.all.divBI.innerHTML = "<input type='hidden' name='iBI' value='"+iBidx+"'><input type='text' name='sBNo' value='"+sBNo+"'onclick='jsSetBI();' size='12'>"
 	self.close();
 }
 
 //수정
 function jsMod(sBidx){
 	location.href= "regBusiness.asp?iBI="+sBidx;
 }
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" >  
<tr>
	<td>업체정보<br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frmS" method="get" action="popBusiness.asp"> 
		<tr align="center" bgcolor="#FFFFFF" >
			<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
			<td align="left">
				 사업자등록번호 : <input type="text" name="sBNo" value="<%=sBusiNo%>" size="12">(-포함)&nbsp;&nbsp;
				 업체명 : <input type="text" name="sBNa" value="<%=sBusiName%>" size="20"> 
			</td>
			<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
			</td> 
		</tr>
		</form>
		</table>	
	</td>
</tr>
<tr>
	<td><a href="regBusiness.asp">+신규등록</a></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center"> 
			<td>Idx</td>
			<td>사업자등록번호</td>
			<td>업체명</td>
			<td>담당자</td>
			<td>선택</td>
		</tr>
		<%IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF" align="center">
			<td><a href="javascript:jsMod('<%=arrList(0,intLoop)%>');"><%=arrList(0,intLoop)%></a></td>
			<td><%=arrList(1,intLoop)%></td>
			<td><%=arrList(2,intLoop)%></td>
			<td><%=arrList(3,intLoop)%></td>
			<td><a href="javascript:jsChoice('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>');">[선택]</a></a></td>
		</tr>
	<%		Next	
		ELSE%>
		<tr bgcolor="#FFFFFF" align="center">
			<td colspan="5">등록된 내용이 없습니다.</td>
		<%END IF%>
		</table>
	</td>
</tr>
</table>
</body>
</html>
 <!-- #include virtual="/lib/db/dbclose.asp" --> 