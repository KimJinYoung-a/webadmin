<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  당첨자 등록
' History : 2008.04.11 정윤정 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/designfingersCls.asp"-->
<%
'--------------------------------------------------------
' 변수선언
'--------------------------------------------------------
Dim iDFSeq
Dim clsDF
Dim iDFType, sTitle, dPrizeDate,sDFTypeDesc
Dim eCode,egKindCode
 
 iDFSeq  = requestCheckVar(request("iDFS"),10)
 eCode = 1				'디자인핑거스 이벤트 번호
 egKindCode = iDFSeq	'디자인핑거스회차
'--------------------------------------------------------
' 이벤트 데이터 가져오기
'--------------------------------------------------------
	 set clsDF = new CDesignFingers
	clsDF.FDFSeq = iDFSeq		
	clsDF.fnGetDFSummary	
	
	iDFType 	 = clsDF.FDFType 	
	sTitle  	 = clsDF.FTitle  		
	dPrizeDate   = clsDF.FPrizeDate 
	sDFTypeDesc  = clsDF.FDFTypeDesc
	set clsDF = nothing
	
	
%>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="0" >
	<tR>
		<td>		
		<span style="height:25px;padding:10 0 5 0"><img src="/images/icon_arrow_link.gif" align="absmiddle"> 당첨관리 : 한번 등록된 당첨자는 취소할 수 없습니다. 입력시 주의해 주세요</span><br>		
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">핑거스ID</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=iDFSeq%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">구분</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sDFTypeDesc%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">제목</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=sTitle%></td>					 
					<td width="80" align="center"  bgcolor="<%= adminColor("tabletop") %>">당첨발표일</td>
					<td width="100" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=dPrizeDate%></td>					 
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<!-- 당첨자 등록-->		
		<!-- #include virtual="/admin/eventmanage/common/inc_eventprize.asp"-->	
		<!-- /당첨자 등록-->
	</td>
</tr>	
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->