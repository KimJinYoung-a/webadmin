<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
 Dim clsDF,clsDFCode
 Dim arrList, intLoop
 Dim iDFSeq, sTitle
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim vItemID, vUserID
  
 vItemID = requestCheckVar(Request("itemid"),10)
 vUserID = requestCheckVar(Request("userid"),50)		'제목
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 
	IF iCurrpage = "" THEN	iCurrpage = 1
	iPageSize = 10		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
'//리스트 가져오기	
 set clsDF = new CDesignFingers
 	clsDF.FCPage = iCurrpage	'현재페이지
	clsDF.FPSize = iPageSize '한페이지에 보이는 레코드갯수
	clsDF.FItemID = vItemID
	clsDF.FUserid = vUserID
 	arrList = clsDF.fnGetRecommendList
 	iTotCnt = clsDF.FTotCnt	'전체 데이터  수
 set clsDF = nothing
 

 iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
 IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1
 	
%>
<script language="javascript">
<!--
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	function jsGoPage(iP){
		document.frmPage.iC.value = iP;
		document.frmPage.submit();
	}
//-->
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" >
<form name="frmFile" method="post">
<input type="hidden" name="iDFS" value="">
</form>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">			
			<form name="frmSearch" method="get" action="recommend_list.asp">	
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
				<td align="left">
					Item ID: <input type="text" name="itemid" value="<%= vItemID %>" size="10">
					&nbsp;회원ID:<input type="text" name="userid" value="<%= vUserID %>" size="10">
				</td>
				<td  width="110" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
					<input type="button" class="button_s" value="전체보기" onClick="location.href='recommend_list.asp';">
				</td>
			</tr>
			</form>	
		</table>
	</td>	
</tr>
<tr>
	<td height="30"></td>
</tr>
<tr>
	<td> 
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a"  bgcolor="#CCCCCC">		
		<tr bgcolor="#EFEFEF">
			<td width="40" align="center" nowrap>IDX</td>
			<td width="120" align="center" nowrap>작성자</td>
			<td width="90" align="center" nowrap>ItemID</td>
			<td align="center">내용</td>
			<td width="70" align="center" nowrap>등록일</td>
			<td width="80" align="center" nowrap>관리</td>			
		</tr>
		<%IF isArray(arrList) THEN%>
		<% For intLoop =0 To UBound(arrList,2) %>	
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="center"><%=arrList(7,intLoop)%>(<%=arrList(1,intLoop)%>)</td>
			<td align="center"><img src="http://webimage.10x10.co.kr/image/small/<%=Num2Str(CStr(Clng(arrList(2,intLoop)) \ 10000),2,"0","R")%>/<%=arrList(6,intLoop)%>"><%=arrList(2,intLoop)%></td>
			<td align="center" ><%=Replace(arrList(3,intLoop),vbCrLf,"<br>")%></td>
			<td align="center"><%=FormatDate(arrList(5,intLoop),"0000.00.00")%></td>
			<td align="center">
			<% IF arrList(4,intLoop) = "Y" Then %>
			<input type="button" value="삭제하기" class="button" onClick="location.href='recommend_proc.asp?gubun=d&idx=<%=arrList(0,intLoop)%>&useyn=N';"></td>
			<% Else %>
			<input type="button" value="되살리기" class="button" onClick="location.href='recommend_proc.asp?gubun=d&idx=<%=arrList(0,intLoop)%>&useyn=Y';"></td>
			<% End If %>
		</tr> 
		<% Next%>
		<%ELSE%>
		<tr bgcolor="#FFFFFF">
			<td colspan="8" align="center">등록된 내역이 없습니다.</td>
		</tr>
		<%END IF%>	
		</table>
	</td>		
</tr>
<tr>
	<td>
		<!-- 페이징처리 -->
<%
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1

If (iCurrpage mod iPerCnt) = 0 Then
	iEndPage = iCurrpage
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<form name="frmPage" method="get" action="recommend_list.asp">
		<input type="hidden" name="menupos" value="<%= menupos %>">
		<input type="hidden" name="iC" value="">
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
	</td>
</tr>
</table>	

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->