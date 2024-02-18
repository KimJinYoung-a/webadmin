<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<%
'###########################################################
' Description : 전자결제
' Hieditor : 정윤정 생성
'			 2022.07.11 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%   
Dim reportidx, payrequestidx,adminID,Comment,sRectAuthId
Dim arrComm, intC
dim clseapp 
 
reportidx = requestCheckvar(Request("iridx"),10)
payrequestidx = requestCheckvar(Request("ipridx"),10)
sRectAuthId = requestCheckvar(Request("sRAId"),32)    
set clseapp = new CEApproval
	clseapp.Freportidx 		= reportidx 
	clseapp.Fpayrequestidx= payrequestidx 
	arrComm			= clseapp.fnGetCommentList
set clseapp = nothing  
 
%>		 
<%IF isArray(arrComm) THEN  
	For intC = 0 To UBound(arrComm,2)
	%>  
	 <span style="font-size:11px;color:#696969"><%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%></span>&nbsp;<%IF  sRectAuthId = arrComm(2,intC) THEN%><a href="javascript:jsDelCmt(<%=arrComm(0,intC)%>);"><img src="http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif" border="0"></a><%END IF%>
	<br>
	<div style="padding:5px;border-bottom:1px solid #BABABA;width=100%"><%= ReplaceBracket(arrComm(1,intC)) %></div><Br>
<%	Next
END IF%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->