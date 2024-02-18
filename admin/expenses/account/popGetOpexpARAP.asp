<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 수지항목 리스트 - 공통사용
' History : 2011.11.15 정윤정  생성
'	jsSetARAP 스크립트 함수 opener에서 생성해서 선택처리
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpArapCls.asp"-->
<!-- #include virtual="/lib/classes/linkedERP/arapCls.asp"-->
<%
Dim clsARAP
Dim arrList, intLoop
Dim sARAP_GB,sCASH_FLOW,sARAP_NM

sARAP_GB = requestCheckvar(Request("rdoGB"),3)  
sCASH_FLOW = requestCheckvar(Request("selFlow"),3)  
sARAP_NM = requestCheckvar(Request("sNM"),50)   

Set clsARAP = new COpExpAccount
	 clsARAP.FARAP_GB		=sARAP_GB 	
	 clsARAP.FCASH_FLOW =sCASH_FLOW 
	 clsARAP.FARAP_NM   =sARAP_NM 	
	arrList = clsARAP.fnGetArapOutList 
Set clsARAP = nothing
%>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF"> 
<tr>
	<td><strong>수지항목  선택</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="post" action="popGetOpExpARAP.asp">  
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">
					 구분:
						<input type="radio" name="rdoGB" value=""<%IF sARAP_GB="" THEN%>checked<%END IF%>>전체
						<input type="radio" name="rdoGB" value="1" <%IF sARAP_GB="1" THEN%>checked<%END IF%>>수입
						<input type="radio" name="rdoGB" value="2" <%IF sARAP_GB="2" THEN%>checked<%END IF%>>지출
						&nbsp; &nbsp; &nbsp;
						분류:
						<select name="selFlow">
							<option value="">전체</option>
							<option value="001"  <%IF sCASH_FLOW="001" THEN%>selected<%END IF%>>영업</option>
							<option value="002"  <%IF sCASH_FLOW="002" THEN%>selected<%END IF%>>투자</option>
							<option value="003"  <%IF sCASH_FLOW="003" THEN%>selected<%END IF%>>재무</option>
						</select>
					</td>
					
						
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>수지항목명: <input type="text" name="sNM" value="<%=sARAP_NM%>" size="20">
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
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<td>코드</td> 
		 	<td>구분</td>  
			<td>분류</td>  
			<td>수지항목</td>  
			<td>연결계정과목</td>  
			<td>매입/매출거래종류</td> 
			<td>선택</td>  
		</tr> 
		<%IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
			%>
		<tr bgcolor="#FFFFFF"  align="center">
			<td><%=arrList(0,intLoop)%></td>
		 	<td><%=fnGetARAP_GB(arrList(1,intLoop))%></td> 
		 	<td><%=fnGetARAP_Cash(arrList(3,intLoop))%></td> 
		 	<td><%=arrList(2,intLoop)%></td> 
		 	<td align="left">[<%=arrList(4,intLoop)%>] <%=arrList(5,intLoop)%></td>  
		 	<td><%=arrList(7,intLoop)%></td> 
		 	<td><input type="button" class="button" value="선택" onClick="opener.jsSetARAP('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>');self.close();"> </td>
		</tr>  
	<%	Next %>
	<%ELSE%>
	<tr bgcolor="#FFFFFF"  align="center">
			<td colspan="7" align="Center">등록된 내용이 없습니다.</td>
		</tr>
	<% 	END IF%>
		</table>	
	</td> 
</tr>  
</table>
<!-- 페이지 끝 -->
</body>
</html>
 