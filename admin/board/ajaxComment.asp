<%@language="VBScript" %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/commentCls.asp"--> 
<%   
Dim iboard_idx,  adminID,Comment,sRectAuthId
Dim arrComm, intC
dim clscomm 
 
iboard_idx = requestCheckvar(Request("ibidx"),10) 
sRectAuthId = requestCheckvar(Request("sRAId"),32)    
set clscomm = new CComment
	clscomm.FboardIdx = iboard_idx
	arrcomm = clscomm.fnGetCommentList
set clscomm = nothing
 
%>		 
<%IF isArray(arrComm) THEN  
	For intC = 0 To UBound(arrComm,2)
	%>  
	 <span style="font-size:11px;color:#696969"><%=arrComm(4,intC)%>(<%=arrComm(2,intC)%>)&nbsp;<%=formatdate(arrComm(3,intC),"0000.00.00")%></span>&nbsp;<%IF  sRectAuthId = arrComm(2,intC) THEN%><a href="javascript:jsDelCmt(<%=arrComm(0,intC)%>);"><img src="http://fiximage.10x10.co.kr/web2009/common/cmt_del.gif" border="0"></a><%END IF%>
	<br>
	<div style="padding:5px;border-bottom:1px solid #BABABA;width=100%"><%=arrComm(1,intC)%></div><Br>
<%	Next
END IF%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->
 