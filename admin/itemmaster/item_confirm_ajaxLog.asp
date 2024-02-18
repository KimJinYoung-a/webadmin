<%@ language=vbscript %>
<%
	Option Explicit
	Response.Expires = -1440
%>
<% response.Charset="euc-kr" %> 
<%
'########################################################### 
' Description :  승인대기상품 상세리스트 진행일자
' History : 2014.01.06 정윤정 생성
'						currstate: 0-승인반려,1-승인대기,2-승인보류,5-승인대기(재요청),7-승인완료,9-업체취소
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"--> 
<%
Dim clsWait, itemid ,arrlist, intLoop
itemid =  requestCheckvar(Request("itemid"),16)
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog
 set clsWait = nothing
%>
 <%IF isArray(arrList) THEN %>
<table border=0 cellpadding=5 cellspacing=5 bgcolor="#EFEFEF" class="a"> 
	<tr>
		<td>임시코드: [<%=itemid%>]<hr width="100%">
			<%
				For intLoop = 0 To UBound(arrList,2)
			%>
			<font color="<%=GetCurrStateColor(arrList(2,intLoop))%>"><%=fnGetCurrStateShortName(arrList(2,intLoop))%></font>: <%=arrList(4,intLoop)%><br>
			<%		
				Next
			%> 
		</td>
	</tr>
</table>
<%END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->