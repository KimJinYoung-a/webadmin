<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/partnerFaqCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
 dim clsFaq, arrFaq, intLoop
 dim iCurrpage, iPageSize, iPercnt,iTotCnt,iTotalPage
 dim stType, stTitle, stContents, stUsername
 dim selSearch, strSearch
 dim strParm
 
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 stType 		= requestCheckVar(Request("selFaqT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	 strParm = "iC="&iCurrpage&"&stfaqT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
	 
	if selSearch="1" then
		stTitle			= strSearch
		stContents	= ""
		stUsername	= ""
	elseif selSearch="2" then
		stTitle			= ""
		stContents	= strSearch
		stUsername	= ""
	elseif  selSearch="3" then
		stTitle			= ""
		stContents	= ""
		stUsername	= strSearch
	else
		stTitle			= ""
		stContents	= ""
		stUsername	= ""	
	end if		
	
	set clsFaq = new CFaq
	clsFaq.FRectType = stType
	clsFaq.FRectTitle = stTitle
	clsFaq.FRectConts = stContents
	clsFaq.FRectUserName = stUsername
	clsFaq.FPSize = iPageSize
	clsFaq.FCPage = iCurrpage
	arrFaq = clsFaq.fnGetFaqList
	iTotCnt = clsFaq.FTotCnt
	set clsFaq = nothing
		
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수	
%>
<script type="text/javascript"> 
	function jsNewReg(){
		location.href="partnerfaqReg.asp?menupos=<%=menupos%>";
	}
	
	function jsSearch(){
		document.frmSearch.submit();
	}
</script>
<!-- 표 상단바 시작-->
<form name="frmSearch" method="post" action="/admin/board/partnerfaqList.asp">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
				<tr>
					<td>FAQ 구분: &nbsp;<% fnOptFaqType stType %> 
						&nbsp;&nbsp;	&nbsp;&nbsp;
						<select name="selSearch" class="select">
							<option value="">--선택--</option>
							<option value="1" <%if selSearch="1" then%>selected<%end if%>>제목</option>
							<option value="2" <%if selSearch="2" then%>selected<%end if%>>내용</option>
							<option value="3" <%if selSearch="3" then%>selected<%end if%>>등록자</option>
							</select>
							<input type="text" size="30" name="strSearch" value="<%=strSearch%>">
					</td>
				</tr>
				</table>
		</td>
    <td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:jsSearch();">
		</td> 
	</tr>
</table>
</form>
<!-- 표 상단바 끝-->
<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >
    <tr height="40" valign="bottom">
    	<td align="right">
        	<input type="button" value="신규등록" onclick="jsNewReg();" class="button">
	    </td> 
	</tr>
</table>
<!-- 표 중간바 끝-->
<!--본문내용 시작-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="19">검색결과 : <b><%=iTotCnt%></b>&nbsp;&nbsp;페이지 : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>Idx</td>
    	<td>FAQ 구분</td>
    	<td>제목</td>
    	<td>등록자</td>
    	<td>등록일</td> 
    </tr>
    <%IF isArray(arrFaq) THEN
    		For intLoop = 0 To UBound(arrFaq,2) 
    	%>
    <tr align="center" bgcolor="#FFFFFF" height="25">
    	<td><a href="/admin/board/partnerfaqView.asp?fidx=<%=arrFaq(0,intLoop)%>&menupos=<%=menupos%>&<%=strParm%>"><%=arrFaq(0,intLoop)%></a></td>
    	<td><%=fnDispFaqType(arrFaq(1,intLoop))%></td>
    	<td><a href="/admin/board/partnerfaqView.asp?fidx=<%=arrFaq(0,intLoop)%>&menupos=<%=menupos%>&<%=strParm%>"><%=arrFaq(2,intLoop)%></a></td>
    	<td><%=arrFaq(6,intLoop)%></td>
    	<td><%=arrFaq(5,intLoop)%></td>
    </tr>
  <%	Next
    ELSE%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="5">등록된 내용이 없습니다.</td>
    </tr>
    <%END IF%>
 </table> 
 <!--본문내용 끝-->
<!-- 페이징처리 --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
