<%
	Dim iStartPage,iEndPage,iX,iPerCnt
	Dim ijScript
%>	<!-- ####### [1] 전체문서 //-->
	<div id="divForm1" style="display:<%IF iFormType <> 1 THEN%>none<%END IF%>;">
	<% IF iFormType ="1" THEN
		icateidx1 = requestCheckvar(Request("selC1"),10)
		icateidx2 = requestCheckvar(Request("iC2"),10)
		sEdmsName = requestCheckvar(Request("sEN"),128)
		if icateidx1 = "" then icateidx1 = 0
		if icateidx2 = "" then icateidx2= 0

		Set clsedms = new Cedms
			clsedms.Fcateidx1 = icateidx1
			clsedms.Fcateidx2	= icateidx2
			clsedms.Fedmsname	= sEdmsName
			clsedms.FCurrPage 	= iCurrPage
			clsedms.FPageSize 	= iPageSize
			arrList = clsedms.fnGetEappEdmsList
			iTotCnt = clsedms.FTotCnt
			iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	%>
	<table width="100%"  cellpadding="0" cellspacing="1" class="a" border="0">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
				<td align="left">
					대카테고리 :
					<select name="selC1" id="selC1">
					<option value="0">전체</option>
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select>
					&nbsp;&nbsp;중카테고리 :
					<span id="divCL">
					<select name="selC2" id="selC2">
					<option value="0">전체</option>
				<% 	IF icateidx1 > 0 THEN	'대카테고리 선택 후 중카테고리 선택가능하게
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2
					END IF
				%>
					</select>
					</span>&nbsp;&nbsp;문서명 : <input type="text" name="sEN" value="<%=sEdmsName%>" size="20">
				</td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="jsSearch();"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td><br>검색결과 : <b><%=iTotCnt%></b> &nbsp;페이지 : <b><%= iCurrPage %> / <%=iTotalPage%></b>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("gray") %>">
				<!-- <td>idx</td> -->
				<td>문서코드</td>
				<td>카테고리</td>
				<td>문서명</td>
				<td>최종결재자</td>
				<td>결제요청서</td>
			</tr>
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
			%>
					<tr height=30 align="center" bgcolor="#FFFFFF">
						<!-- <td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%=arrList(0,intLoop)%></td> -->
						<td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%=arrList(7,intLoop)%></td>
						<td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%=arrList(2,intLoop)%> > <%=arrList(4,intLoop)%></td>
						<td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%=arrList(6,intLoop)%></td>
						<td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%=arrList(16,intLoop)%></td>
						<td><a href="javascript:jsSetDoc(<%=arrList(0,intLoop)%>,'<%=arrList(17,intLoop)%>');"><%IF arrList(17,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
					</tr>
			<%
					Next
				ELSE
			%>
					<tr height=30 align="center" bgcolor="#FFFFFF"><td colspan="8">등록된 내용이 없습니다.</td></tr>
			<%	END IF %>
			</table>
		</td>
	</tr><!-- 페이지 시작 -->
	<!-- #include virtual="/admin/approval/eapp/include_regeappform_list_paging.asp" -->
	</table>
	<% END IF %>
	</div>


<!-- ####### [2] 결제관련문서 //-->
	<div id="divForm2" style="display:<%IF iFormType <> 2 THEN%>none<%END IF%>;">
	<%
	IF iFormType ="2" THEN
	sARAPNM =  requestCheckvar(Request("sANM"),50)
 	sedmsNM =  requestCheckvar(Request("sENM"),60)

	Set clsALE = new CArapLinkEdms
		clsALE.FARAP_NM 	= sARAPNM
		clsALE.FEdmsName 	= sedmsNM
		clsALE.FCurrPage	= iCurrPage
		clsALE.FPageSize	= iPageSize
		arrList = clsALE.fnGetEappArapLinkNPayEdmsList
		iTotCnt = clsALE.FtotCnt
	Set clsALE = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	%>
	<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">수지항목: <input type="text" name="sANM" value="<%=sARAPNM%>" size="20">&nbsp;&nbsp;문서명: <input type="text" name="sENM" value="<%=sedmsNM%>" size="20"></td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="jsSearch();"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td><br>검색결과 : <b><%=iTotCnt%></b> &nbsp;
			<!-- 상단 띠 시작 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<!-- <td>idx</td> -->
				<td>문서코드</td>
				<td>카테고리</td>
				<td>문서명</td>
				<td>최종결재자</td>
				<td>결제요청서</td>
				<td>수지항목</font></td>
				<td>연결계정과목</font></td>
			</tr>
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
					if isNULL(arrList(1,intLoop)) then
						ijScript = "javascript:jsSetDoc('"&arrList(2,intLoop)&"','"&arrList(13,intLoop)&"');"
					else
					    ijScript = "javascript:jsSelectEApp('"&arrList(1,intLoop)&"','"&arrList(2,intLoop)&"');"
				    end if
			%>
					<tr height=30 align="center" bgcolor="#FFFFFF">
						<!-- <td><a href="<%=ijScript%>"><%=arrList(0,intLoop)%></td>-->
						<td nowrap><a href="<%=ijScript%>"><%=arrList(6,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(9,intLoop)%> > <%=arrList(10,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(5,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(12,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%IF arrList(13,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
						<td align="left"><a href="<%=ijScript%>"><% if isNULL(arrList(1,intLoop)) then %><font color="gray">결제요청시 선택</font><% else %><font color="blue">[<%=arrList(1,intLoop)%>] <%=arrList(3,intLoop)%></font><% end if %></a></td>
						<td align="left"><a href="<%=ijScript%>"><% if isNULL(arrList(1,intLoop)) then %><font color="gray">결제요청시 선택</font><% else %><font color="blue">[<%=arrList(14,intLoop)%>] <%=arrList(4,intLoop)%></font><% end if %></a></td>
					</tr>
			<%
					Next
				ELSE
			%>
					<tr height=5 align="center" bgcolor="#FFFFFF"><td colspan="10">등록된 내용이 없습니다.</td></tr>
			<% END IF %>
			</table>
		</td>
	</tr><!-- 페이지 시작 -->
	<!-- #include virtual="/admin/approval/eapp/include_regeappform_list_paging.asp" -->
	</table>
	<%END IF%>
	</div>


<!-- ####### [3] 최근사용문서 //-->
	<div id="divForm3" style="display:<%IF iFormType <> 3 THEN%>none<%END IF%>;">
	<%
	IF iFormType ="3" THEN
	sARAPNM =  requestCheckvar(Request("sANM"),50)
 	sedmsNM =  requestCheckvar(Request("sENM"),60)

	Set clsALE = new CArapLinkEdms
		clsALE.FARAP_NM 	= sARAPNM
		clsALE.FEdmsName 	= sedmsNM
		clsALE.FadminId		= session("ssBctId")
		clsALE.FCurrPage	= iCurrPage
		clsALE.FPageSize	= iPageSize
		arrList = clsALE.fnGetEappArapLinkNPayEdmsList
		iTotCnt = clsALE.FtotCnt
	Set clsALE = nothing
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
	%>
	<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">검색조건</td>
				<td align="left">수지항목: <input type="text" name="sANM" value="<%=sARAPNM%>" size="20">&nbsp;&nbsp;문서명: <input type="text" name="sENM" value="<%=sedmsNM%>" size="20"></td>
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="jsSearch();"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td><br>검색결과 : <b><%=iTotCnt%></b> &nbsp;
			<!-- 상단 띠 시작 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<!-- <td>idx</td> -->
				<td>문서코드</td>
				<td>카테고리</td>
				<td>문서명</td>
				<td>최종결재자</td>
				<td>결제요청서</td>
				<td>수지항목</font></td>
				<td>연결계정과목</font></td>
			</tr>
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2)
					if isNULL(arrList(1,intLoop)) then
						ijScript = "javascript:jsSetDoc('"&arrList(2,intLoop)&"','"&arrList(13,intLoop)&"');"
					else
					    ijScript = "javascript:jsSelectEApp('"&arrList(1,intLoop)&"','"&arrList(2,intLoop)&"');"
				    end if
			%>
					<tr height=30 align="center" bgcolor="#FFFFFF">
						<!-- <td><a href="<%=ijScript%>"><%=arrList(0,intLoop)%></td> -->
						<td nowrap><a href="<%=ijScript%>"><%=arrList(6,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(9,intLoop)%> > <%=arrList(10,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(5,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%=arrList(12,intLoop)%></td>
						<td><a href="<%=ijScript%>"><%IF arrList(13,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>

						<td align="left"><a href="<%=ijScript%>"><% if isNULL(arrList(1,intLoop)) then %><font color="gray">결제요청시 선택</font><% else %><font color="blue">[<%=arrList(1,intLoop)%>] <%=arrList(3,intLoop)%></font><% end if %></a></td>
						<td align="left"><a href="<%=ijScript%>"><% if isNULL(arrList(1,intLoop)) then %><font color="gray">결제요청시 선택</font><% else %><font color="blue">[<%=arrList(14,intLoop)%>] <%=arrList(4,intLoop)%></font><% end if %></a></td>
					</tr>
			<%
					Next
				ELSE
			%>
					<tr height=5 align="center" bgcolor="#FFFFFF"><td colspan="10">등록된 내용이 없습니다.</td></tr>
			<% END IF %>
			</table>
		</td>
	</tr><!-- 페이지 시작 -->
	<!-- #include virtual="/admin/approval/eapp/include_regeappform_list_paging.asp" -->
	</table>
	<%END IF%>
	</div>
<%
	If iFormType = "1" THEN
		Set clsedms = nothing
	Else
		Set clsALE = nothing
	End If
%>