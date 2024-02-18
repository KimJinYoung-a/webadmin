<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/artistboard_cls.asp"-->
<%
Dim CBList, arrList, intLoop,i
dim lecuserid, iCurrpage,iPageSize, iTotCnt, iSearchType, sSearchTxt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
		
	iCurrpage = requestCheckVar(Request("iC"),10)
	iTotCnt = requestCheckVar(Request("iTC"),10)	
	iSearchType =requestCheckVar(Request("iSType"),1)
	sSearchTxt =requestCheckVar(Request("sSTxt"),50)
	
	IF iCurrpage = "" THEN iCurrpage = 1
	iPageSize = 20
	iPerCnt = 10	
	
	IF iTotCnt = "" THEN iTotCnt = -1
	IF iSearchType = "" THEN iSearchType = -1
			
	Set CBList = new CArtistRoomBoard
		
		CBList.FLecuserid = lecuserid
		CBList.FSearchType = iSearchType
		CBList.FSearchTxt = sSearchTxt
		CBList.FCPage = iCurrpage
		CBList.FPSize = iPageSize
		CBList.FTotCnt = iTotCnt
		arrList = CBList.fnGetList
		iTotCnt = CBList.FTotCnt
		
	Set CBList = nothing
'전체 페이지 수
	iTotalPage 	=  Int(iTotCnt/iPageSize)	
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1	
%>
<script language='javascript'>
<!--
	function chk_form()
	{
		var frm = document.frm;

		if(!frm.iSType.value)
		{
			alert("검색 조건을 선택해주십시오.");
			frm.iSType.focus();
			return false;
		}
		else if(!frm.sSTxt.value)
		{
			alert("검색어를 입력해주십시오.");
			frm.sSTxt.focus();
			return false;
		}

		return;
	}

	function jsGoPage(iP){
		document.frm.iC.value = iP;
		document.frm.submit();	
	}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method="POST" action="artistboard_list.asp" onSubmit="return chk_form()">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
		 검색
		<select name="iSType">
			<option value="" >선택</option>			
			<option value="0" <%IF iSearchType=0 THEN%>selected<%END IF%>>강사명</option>
			<option value="1" <%IF iSearchType=1 THEN%>selected<%END IF%>>아이디</option>
			<option value="2" <%IF iSearchType=2 THEN%>selected<%END IF%>>제목</option>
		</select>		
		<input type="text" name="sSTxt" size="20" value="<%= sSearchTxt %>">
       	<input type="image" src="/admin/images/search2.gif" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>

</table>
<!-- 상단 검색폼 끝 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="6" align="left">검색건수 : <%= iTotCnt%> 건 Page : <%= iCurrpage %>/<%=iTotalPage%></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">번호</td>		
		<td align="center" width="100">강사명</td>
		<td align="center" width="100">등록자</td>
		<td align="center">제목</td>		
		<td align="center" width="50">조회수</td>
		<td align="center" width="80">등록일</td>
	</tr>
	<%IF isArray(arrList) THEN
        For intLoop = 0 To UBound(arrList,2)	
     %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=iTotCnt-intLoop-(iPageSize*(iCurrpage-1))%></td>		
		<td><%=arrList(3,intLoop)%></td>
		<td><%=arrList(4,intLoop)%></td>
		<td align="left">&nbsp;
		<%IF arrList(2,intLoop) > 0 THEN%>
		<%For i = 0 To arrList(2,intLoop) %>			
		&nbsp;
		<%Next%>
		<img src="http://www.thefingers.co.kr/images/artistroom/10.gif" align="absmiddle">
		<%END IF%>
		<a href="artistboard_view.asp?idx=<%=arrList(0,intLoop)%>&lecuserid=<%=arrList(3,intLoop)%>"><%=db2html(arrList(5,intLoop))%></a>
		<%IF DateDiff("d",arrList(6,intLoop),now()) < 7 THEN%><img src="http://www.thefingers.co.kr/images/artistroom/new.gif" width="15" height="14"><%END IF%>
		</td>		
		<td><%=arrList(7,intLoop)%></td>
		<td><%= FormatDate(arrList(6,intLoop),"0000.00.00") %></td>
	</tr>
	<%
		next
	 ELSE	
	%>
	<tr align="center" bgcolor="#FFFFFF"><td colspan="6" align="center">등록된 내용이 없습니다.</td></tr>
	<%	
	  END IF	
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
						<!-- 페이징처리 -->
			<%		
            iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
			
			If (iCurrpage mod iPerCnt) = 0 Then																
				iEndPage = iCurrpage
			Else								
				iEndPage = iStartPage + (iPerCnt-1)
			End If	
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" >
			    <tr valign="bottom" height="25">			        
			        <td valign="bottom" align="center">
			        <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>			     
			    </tr>			    
			   </form> 
			</table>
			<!-- /페이징처리 -->
				</td>
				<td width="80" align="right">
					<a href="artistboard_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
