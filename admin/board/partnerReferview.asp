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
<!-- #include virtual="/lib/classes/board/partnerReferCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
 dim clsref, arrref, intLoop
 dim iCurrpage, iPageSize, iPercnt,iTotCnt,iTotalPage
 dim sTitle, tContents, sType, sregid, sregname, dregdate
 dim refidx
 Dim sMode
 dim arrFile ,intF
 dim strParm
 dim stType, selSearch,strSearch
 
 '--리스트 검색 파라미터================================================================
 iCurrpage = requestCheckVar(Request("iC"),10)	'현재 페이지 번호 
 stType 		= requestCheckVar(Request("strefT"),4)
 selSearch = requestCheckVar(Request("selSearch"),10)
 strSearch = requestCheckVar(Request("strSearch"),200)
  
  strParm = "iC="&iCurrpage&"&strefT="&stType&"&selSearch="&selSearch&"&strSearch="&strSearch
 '--================================================================
  
 refidx =  requestCheckVar(Request("fidx"),10)
 sMode ="I"
 
 if refidx <> "" THEN	 
 		sMode ="U"
		set clsref = new CRefer
	 	clsref.FrefIdx = refidx
	 	clsref.FnGetReferConts
	 	sType 		= clsref.FrefType
	 	sTitle 		= clsref.FTitle
	 	tContents = clsref.FContents
	 	sregid 		= clsref.Fregid
	 	sregname 	= clsref.Fregname
	 	dregdate	= clsref.Fregdate
	 	
	 	arrFile   = clsref.fnGetAttachFile
		set clsref = nothing
	
	END IF	
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
  
<script type="text/javascript"> 
	 
	
	function jsCancel(){
		location.href="/admin/board/partnerReferList.asp?menupos=<%=menupos%>&<%=strParm%>"
	}
	
	  

//파일 다운로드
    function jsDownload(sDownURL, sRFN, sFN){
    var winFD = window.open(sDownURL+"/linkweb/board/procDownload.asp?sRFN="+sRFN+"&sFN="+sFN,"popFD","");
    winFD.focus();
 }
</script>
<form name="frm" method="post" action="partnerReferProc.asp?<%=strParm%>"> 
	<input type="hidden" name="hidM" value="<%=sMode%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="850" border="0" class="a" cellpadding="3" cellspacing="0" > 
<tr>
	<td >
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<%if refidx <> "" THEN	 %>
			<tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">idx</td>
		   		<td bgcolor="#FFFFFF"><%=refidx%><input type="hidden" name="fidx" value="<%=refidx%>"></td>
		   	</tr>
			 <tr>
			<%end if%> 	
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">구분</td>
		   		<td bgcolor="#FFFFFF"><%fnOptReferType sType%></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">제목</td>
		   		<td bgcolor="#FFFFFF"><input type="text" name="sT" size="60" maxlength="60" value="<%=sTitle%>"></td>
		   	</tr>
		   	 <tr>
		   		<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">내용</td>
		   		<td bgcolor="#FFFFFF">
		   			 <%=tContents%> 
		   		</td>
		   	</tr>
		   	<tr>
					<td align="center" width="100" bgcolor="<%= adminColor("tabletop") %>">첨부파일</td>
					<td bgcolor="#FFFFFF"> 
					<div id="dFile">
						<% Dim arrFName,arrF, sFName, intF2,intF3, iCount 
						IF isArray(arrFile) THEN
						For intF=0 To UBound(arrFile,2) 
					
								arrF = split(arrFile(2,intF),"/") 
							 	arrFName = arrF(ubound(arrF))
								sFName = split(arrFName,".")(0)  
						%>
						<div id="dF<%=sFName%>"><a href="javascript:jsDownload('<%=uploadImgUrl%>','<%=arrFName%>','<%=arrF(ubound(arrF)-1)&"/"&arrFName%>');"><%=arrFName%></a>&nbsp; 
							<input type="hidden" name="sFileP"   value="<%= arrFile(2,intF)%>"></div>
					<%Next
						END IF
						%> 
						</div> 
					</td>
				</tr>
		 	</table>	 
	</td>
</tr>	 
<tr>
	<td width="100%" align="center" style="padding-top:10px;">
		<input type="button" class="button" value="목록으로" style="width:80px;" onClick="jsCancel();"> &nbsp; 
		<input type="button" class="button" value="수정" style="width:80px;color:red" onClick="location.href='/admin/board/partnerReferReg.asp?fidx=<%=refidx%>&menupos=<%=menupos%>&<%=strParm%>'">
		
	</td>
</tr>
</table>
</form>  	
 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
	session.codePage = 949
%>
