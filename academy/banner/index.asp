<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/academy/lib/academy_function.asp" -->
<!-- #include virtual="/academy/lib/classes/banner_cls.asp" -->
<%
'#############################################
' PageName : /academy/banner/index.asp	
' Description : 핑거스 배너 리스트
' History : 2006.11.16 정윤정 생성
'#############################################

	'// 변수선언
Dim cBanner 
Dim iTotCnt, arrBanner, intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim sLocation, simgUrl
simgUrl = imgFingers & "/contents/banner/"	

	'// 파라미터 값 받기 & 기본 변수 값 세팅 //
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	iPageSize = 20		'한 페이지의 보여지는 레코드 수
	iPerCnt = 10		'보여지는 페이지 간격
	sLocation = Request("selLoc")
	
	'// 데이터 가져오기
 set cBanner = new ClsBanner 	
 	cBanner.FCPage = iCurrpage	
 	cBanner.FPSize = iPageSize	
 	cBanner.FLocation = sLocation
 	arrBanner = cBanner.fnGetBannerList '배너 리스트
 	iTotCnt = cBanner.FBannerCnt	'배너 총 갯수
 set cBanner = nothing
	
	iTotalPage 	=  Int(iTotCnt/iPageSize)	'전체 페이지 수
	IF (iTotCnt MOD iPageSize) > 0 THEN	iTotalPage = iTotalPage + 1
%>
<script language="javascript">
<!--
	function jsGoPage(iP){
		document.frm.iC.value = iP;
		document.frm.submit();	
	}	
	
	function jsPopImg(sImg){
		var winImg;
		winImg = window.open('/academy/lib/popViewImg.asp?sImgUrl='+sImg,'popImg','width=100,height=100,left=10,top=10,scrollbars=1');
		winImg.focus();
	}
	
	function jsDel(iBId){
		if(confirm("삭제하시겠습니까?")){
			document.frmDel.iBId.value = iBId;
			document.frmDel.submit();
		}
	}
//-->
</script>
<form name="frmDel" method="post" action="<%=imgFingers%>/linkweb/processBanner.asp"  enctype="multipart/form-data">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="iBId" value="">
<input type="hidden" name="sType" value="D">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
</form>
<!-- 상단 검색폼 -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method="post" action="index.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="iC" value="">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
		위치구분: 
		<select name="selLoc" onChange="javascript:document.frm.submit();">
		 <option value="">--선택--</option>
		 <%Call sbOptCommCd(sLocation,"I000")%>
		</select>       	
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- /상단 검색폼 -->
<!-- 본문 내용 -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="8" align="left">검색건수 : <%= iTotCnt%> 건 Page : <%= iCurrpage %>/<%=iTotalPage%></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">번호</td>
		<td align="center">이미지</td>
		<td align="center">링크URL</td>
		<td align="center" width="70">위치</td>
		<td align="center" width="50">등록자</td>
		<td align="center" >등록일</td>
		<td align="center" width="50">사용여부</td>
		<td align="center" width="100">처리</td>
	</tr>
	<%IF isArray(arrBanner) THEN%>
		<%FOR intLoop = 0 TO UBound(arrBanner,2)
			'bannerId, imgUrl, linkUrl, commCd, isUsing, regdate, adminID
		%>
	<tr bgcolor="#FFFFFF">
		<td height="30" align="center"><%=iTotCnt-intLoop-(iPageSize*(iCurrpage-1))%></td>
		<td height="30" align="center"><a href="javascript:jsPopImg('<%=simgUrl&arrBanner(3,intLoop)&"/"&arrBanner(1,intLoop)%>');"><img src="<%=simgUrl&arrBanner(3,intLoop)&"/"&arrBanner(1,intLoop)%>" border="0" width="50" height="50"></a></td>
		<td height="30" align="center"><%=arrBanner(2,intLoop)%></td>
		<td height="30" align="center"><%=arrBanner(7,intLoop)%></td>
		<td height="30" align="center"><%=arrBanner(6,intLoop)%></td>
		<td height="30" align="center"><%=arrBanner(5,intLoop)%></td>
		<td height="30" align="center"><%=arrBanner(4,intLoop)%></td>
		<td height="30" align="center">
			<a href="registBanner.asp?iBId=<%=arrBanner(0,intLoop)%>"><img src="/images/icon_modify.gif" border="0"></a>
			<a href="javascript:jsDel(<%=arrBanner(0,intLoop)%>);"><img src="/images/icon_delete.gif" border="0"></a>
		</td>
	</tr>
		<%NEXT%>
	<%ELSE%>	
	<tr bgcolor="#FFFFFF">
		<td colspan="8" height="30" align="center">
			등록된 내용이 없습니다.
		</td>
	</tr>
	<%END IF%>
</table>
<!-- /본문 내용-->
<!-- 페이징처리 -->
<%		
iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	

If (iCurrpage mod iPerCnt) = 0 Then																
	iEndPage = iCurrpage
Else								
	iEndPage = iStartPage + (iPerCnt-1)
End If	

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
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
        <td width="80"><a href="registBanner.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0"></a></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
   </form> 
</table>
<!-- /페이징처리 -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->