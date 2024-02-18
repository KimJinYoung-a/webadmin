<%@ language=vbscript %>
<% option explicit %> 
<%
'###########################################################
' Description : 운영비관리 팀 구분 수정  
' History : 2011.05.30 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<%
Dim sMode
Dim clsPart, clsPartList, arrPartType
Dim  iPartTypeIdx, sPartTypeName, sIsUsing, blnUsing , i
  
iPartTypeIdx = requestCheckvar(Request("hidPT"),10)  
sIsUsing = requestCheckvar(Request("isU"),1)
if sIsUsing="" then sIsUsing="A"

'구분 목록 접수
Set clsPartList = new COpExpPart  
	clsPartList.FIsUsing = sIsUsing
	arrPartType = clsPartList.fnGetOpExpPartTypeListNew
Set clsPartList = Nothing

if isArray(arrPartType) then
	'구분값 가져오기
	if iPartTypeIdx="" then
		iPartTypeIdx = arrPartType(0,0)
	end if

	Set clsPart = new COpExpPart  
			clsPart.FPartTypeIdx  = iPartTypeIdx
			clsPart.fnGetOpExpPartTypeData
			iPartTypeIdx 	= clsPart.FPartTypeIdx 	
			sPartTypeName   = clsPart.FPartTypeName  
			blnUsing 		= clsPart.FIsUsing 	 
	Set clsPart = nothing
End if
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<script type="text/javascript">	
// 이동
function fnMove(i,u) {
	self.location.href = "?hidPT="+i+"&isU="+u;
}

//등록
function jsPartTypeSubmit(){ 
	if( document.frm.sPTN.value==""){
	alert("구분명을 입력해주세요");
	return;
	}
	
	document.frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>운영비사용처 구분 수정</strong><br><hr width="100%"></td>
</tr>
<%
	if isArray(arrPartType) then
%>
<tr>
	<td align="right">
		사용여부 : 
		<select name="isUsing" class="select" onChange="fnMove(<%=iPartTypeIdx%>,this.value)">
		<option value="A" <%=chkIIF(sIsUsing="A","selected","")%>>전체</option>
		<option value="Y" <%=chkIIF(sIsUsing="Y","selected","")%>>사용</option>
		<option value="N" <%=chkIIF(sIsUsing="N","selected","")%>>사용안함</option>
		</select>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<tr align="center">
			<td bgcolor="<%= adminColor("tabletop") %>">구분번호</td>
			<td bgcolor="<%= adminColor("tabletop") %>">구분명</td>
			<td bgcolor="<%= adminColor("tabletop") %>">카드여부</td>
			<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
		</tr>
		<%
			for i=0 to UBound(arrPartType,2)
		%>
		<tr onClick="fnMove(<%=arrPartType(0,i)%>,'<%=sIsUsing%>')" style="cursor:pointer;">
			<td bgcolor="<%=chkIIF(iPartTypeIdx=arrPartType(0,i),"#DFCFFF","#FFFFFF")%>" align="center"><%=arrPartType(0,i)%></td>
			<td bgcolor="<%=chkIIF(iPartTypeIdx=arrPartType(0,i),"#DFCFFF","#FFFFFF")%>"><%=arrPartType(1,i)%></td>
			<td bgcolor="<%=chkIIF(iPartTypeIdx=arrPartType(0,i),"#DFCFFF","#FFFFFF")%>" align="center"><%=chkIIF(arrPartType(3,i),"카드","일반")%></td>
			<td bgcolor="<%=chkIIF(iPartTypeIdx=arrPartType(0,i),"#DFCFFF","#FFFFFF")%>" align="center"><%=chkIIF(arrPartType(2,i),"사용","사용안함")%></td>
		</tr>
		<%
			next
		%>
		</table>
	</td>
</tr>
<%
	end if
%>
<tr>
	<td>
		<!-- 상단 띠 시작 -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
		<form name="frm" method="post" action="procPart.asp"> 
		<input type="hidden" name="hidM" value="T"> 
		<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">  
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">구분번호</td>  
			<td bgcolor="#FFFFFF">
			 <%=iPartTypeIdx%>
			</td>	  
		</tr>
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>" align="center">구분명</td>  
			<td bgcolor="#FFFFFF">
				<input type="text" name="sPTN" size="30" maxlength="60" value="<%=sPartTypeName%>">	
			</td>	  
		</tr> 
		<tr>
		 	<td  bgcolor="<%= adminColor("tabletop") %>"  align="center">사용여부</td>  
			<td bgcolor="#FFFFFF"><input type="radio" name="rdoU" value="1" <%IF blnUsing THEN%>checked<%END IF%>>사용 <input type="radio" value="0"  name="rdoU"  <%IF not blnUsing THEN%>checked<%END IF%>>사용안함</td>
		</tr> 
		</form>
		</table>	
	</td> 
</tr> 	 
<tr>
	<td align="center"><input type="button" value="등록" class="button" onClick="jsPartTypeSubmit();"></td>
</tr>
</table>
<!-- 페이지 끝 -->
</body>
</html>
 



	