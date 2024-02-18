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
 dim sTitle, tContents, sType, sregid, sregname, dregdate
 dim faqidx
 Dim sMode
 dim strParm
 dim stType, selSearch,strSearch
 
 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("stfaqT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&stfaqT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
 '--================================================================

 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호
 faqidx =  requestCheckVar(Request("fidx"),10)
 sMode ="I"
 
 if faqidx <> "" THEN	 
 		sMode ="U"
		set clsFaq = new CFaq
	 	clsFaq.FFaqIdx = faqidx
	 	clsFaq.FnGetFaqConts
	 	sType 		= clsFaq.FFaqType
	 	sTitle 		= clsFaq.FTitle
	 	tContents = clsFaq.FContents
	 	sregid 		= clsFaq.Fregid
	 	sregname 	= clsFaq.Fregname
	 	dregdate	= clsFaq.Fregdate
		set clsFaq = nothing
	
	END IF	
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
  
<script type="text/javascript">  
	 
	function jsCancel(){
		location.href="/admin/board/partnerfaqList.asp?menupos=<%=menupos%>&<%=strParm%>";
	}
	
	 
</script>
<form name="frm" method="post" action="partnerFaqProc.asp?<%=strParm%>"> 
	<input type="hidden" name="hidM" value="<%=sMode%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="850" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<%if faqidx <> "" THEN	 %>
			<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">FAQ idx</td>
		   		<td bgcolor="#FFFFFF"><%=faqidx%><input type="hidden" name="fidx" value="<%=faqidx%>"></td>
		   	</tr>
			 <tr>
			<%end if%> 	
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">FAQ 구분</td>
		   		<td bgcolor="#FFFFFF"><%fnOptFaqType sType%></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">제목</td>
		   		<td bgcolor="#FFFFFF"><%=sTitle%></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">내용</td>
		   		<td bgcolor="#FFFFFF">
		   			 <%=tContents%> 
		   		</td>
		   	</tr>
		 	</table>	 
	</td>
</tr>	
<tr>
	<td width="100%" align="center" style="padding-top:10px;">
		<input type="button" class="button" value="목록으로" style="width:80px;" onClick="jsCancel();"> &nbsp; 
		<input type="button" class="button" value="수정" style="width:80px;color:red" onClick="location.href='partnerfaqReg.asp?fidx=<%=faqidx%>&menupos=<%=menupos%>&<%=strParm%>'">
		
	</td>
</tr>
</table>
</form>  	
 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
