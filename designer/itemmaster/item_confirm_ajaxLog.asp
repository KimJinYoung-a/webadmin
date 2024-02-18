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
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" --> 
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->  
<%
Dim clsWait, itemid ,arrlist, intLoop, sMode
Dim arrOld
itemid =  requestCheckvar(Request("itemid"),16)
sMode	 =  requestCheckvar(Request("hidM"),1)
 set clsWait = new CWaitItemlist2014
 	clsWait.Fitemid = itemid
 	arrList = clsWait.fnGetWaitItemLog 
 	IF not isArray(arrList) THEN
 		arrOld = clsWait.fnGetOldWaitItemLog
	END IF
 set clsWait = nothing
 

 
 IF sMode = "D" THEN '--진행일자
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
<%ELSE '--처리내용%>
 <%IF isArray(arrList) THEN %>
<table border=0 cellpadding=5 cellspacing=5 bgcolor="#EFEFEF" class="a"> 
	<tr>
		<td>임시코드: [<%=itemid%>]<hr width="100%">
			<% Dim intNum
			intNum = 0
				For intLoop = 0 To UBound(arrList,2)
					IF arrList(2,intLoop)="2" THEN
						intNum = intNum+1
			%>
		 	<div style="padding:3"><%=intNum%>차 [<%=arrList(4,intLoop)%>]	<br>
		<%=replace(arrList(3,intLoop),"^","/")%></div>
		<%		END IF	
				Next
			%> 
			*3회 이상 보류 시, 반려처리(재등록불가)
		</td>
	</tr>
</table>
<%ELSEIF isArray(arrOld) THEN %>
<table border=0 cellpadding=5 cellspacing=5 bgcolor="#EFEFEF" class="a"> 
	<tr>
		<td>임시코드: [<%=itemid%>]<hr width="100%">
			<%  IF arrold(4,0)="2" or arrold(4,0)="0" THEN
			%>
		 	<div style="padding:3"> <%=fnGetCurrStateShortName(arrold(4,0))%> : [<%=arrOld(0,intLoop)%>]	<br>
		 <%=arrOld(1,intLoop)%> </div>
		<%		END IF %>
				
			*3회 이상 보류 시, 반려처리(재등록불가)
		</td>
	</tr>
</table>
<%END IF%>
<%END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->