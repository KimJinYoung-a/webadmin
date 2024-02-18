<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
'####################################################
' Description :  업무협조 리스트
' History : 2014.03.07   정윤정 수정
'####################################################
%>
<%
	Dim iTotCnt, arrList,intLoop
	Dim iPageSize, iCurrentpage ,iDelCnt
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDoc_Type, sDoc_Status, sDoc_AnsOX, sSearchMine
	 
	
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sDoc_Status		= NullFillWith(requestCheckVar(Request("doc_status"),10),"x")
	sDoc_Type		= NullFillWith(requestCheckVar(Request("doc_type"),10),"")
	sDoc_AnsOX		= NullFillWith(requestCheckVar(Request("ans_ox"),1),"")
	sSearchMine		= NullFillWith(requestCheckVar(Request("onlymine"),1),"o")
	iPageSize 		= CHKIIF(g_VertiHoriz="h",7,15)
	iPerCnt 		= 10
 
	
	Dim cooperatelist , i
		set cooperatelist = new CCooperate
	 	cooperatelist.FCPage = iCurrentpage
	 	cooperatelist.FPSize = iPageSize
	 	cooperatelist.FDoc_Status = sDoc_Status
	 	cooperatelist.FDoc_Type = sDoc_Type
	 	cooperatelist.FDoc_AnsOX = sDoc_AnsOX
	 	cooperatelist.FDoc_MineOX = sSearchMine 
		arrList = cooperatelist.fnGetMyCooperateList
		iTotCnt = cooperatelist.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<script language="javascript">
function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=410,height=570');
}
function goWrite(didx)
{
	 top.coopDetail.location.href = "cooperate_write.asp?didx="+didx+"";
}
function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}

function mine()
{
	if(!(document.frm.onlyminechk.checked))
	{
		document.frm.onlymine.value = "x";
	}
	else
	{
		document.frm.onlymine.value = "o";
	}
}
function issystem(value)
{
}

function goPopDetail(didx)
{
	 window.open("cooperate_write.asp?ispop=pop&didx="+didx+"","","width=900, height=1000,resizable=yes, scrollbars=yes");
}

//전자결재 품의서 등록 - 업무협조고유번호(scmidx) 
function jsRegEapp(scmidx){ 
	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
	document.frmEapp.iSL.value = scmidx;   
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//전자결재 품의서 내용보기
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/popIndex.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}
</script>
</head>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<div style="height:100%;overflow-y:auto;">
<table width="100%" cellpadding="0" cellspacing="1" class="a" border="0"> 
<tr> 
	<td height="25"><font color="#4E9FC6"><b>보낸업무협조 > <%=NaviTitle(sDoc_Status)%></font></b></font></td>
</tr>
</table>
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="tC" value="">
<input type="hidden" name="ieidx" value="37">  
<input type="hidden" name="iSL" value="">
</form>
<form name="frm" action="my_cooperate.asp" method="get">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
     	요청구분:<%=CommonCode("w","doc_type",sDoc_Type)%>
     	&nbsp;
     	<input type="submit" value="검색" class="button" onfocus="this.blur();">
     	<br>
     	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
     	<% if g_TeamJang="o" or g_PartJang="o" then %>
     	<!--<label id="onlymine"><input type="checkbox" name="onlyminechk" onClick="mine()" value="o" <% If sSearchMine = "o" Then %>checked<% End If %>>나의 작업만 보기</label>//-->
     	<% end if %>
	</td>
</tr> 
</table>
</form>
<p><% If CInt(session("ssAdminLsn")) = 1 AND CInt(session("ssAdminPsn")) = 7 Then %><input type="button" class="button" value="코드관리" onClick="code_manage()"><% End If %><p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#FFFFFF">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="40">NO</td>
		<td>제목</td>
		<td width="80">중요도</td>
		<td width="70">등록일</td>
		<td width="40">답변</td> 
		<td >결재내역</td> 
	</tr>
	<%
		IF isArray(arrList) THEN
			For intLoop =0 To UBound(arrList,2)
	%>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
		<td align="left">
			<font color="silver"><%=CommonCode("v","doc_type",arrList(2,intLoop))%></font>
			<br><span style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=db2html(arrList(1,intLoop))%></span>
			&nbsp;&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2012/category/product_btn_view.png" border="0" style="cursor:pointer" onClick="goPopDetail('<%=arrList(0,intLoop)%>');">
		</td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=CommonCode("v","doc_important",arrList(3,intLoop))%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=FormatDatetime(arrList(6,intLoop),2)%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(7,intLoop)%></td>
		<td nowrap>  <!--'시스템개발 및 수정일때만 결재리스트 보여준다 2014.03.06 정윤정 추가-->
			<%IF (arrList(2,intLoop)="3" )  THEN %>
				<% if isNull(arrList(9,intLoop)) then %>
				<input type="button" class="button"  value="품의서 작성" onClick="jsRegEapp('<%=arrList(0,intLoop)%>');" >
				<% else %>
				<%=fnGetReportState(arrList(10,intLoop))%>&nbsp;
				<input type="button" class="button"   value="품의서 보기" onClick="jsViewEapp('<%=arrList(9,intLoop)%>','<%= arrList(10,intLoop)%>');">
				<% end if%> 
		<%END IF%>
			</td>
	</tr>
	<%
			Next
		Else
	%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
	<%
		End If
	%>
		
	<!-- 페이징처리 -->
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	
	<form name="frmpage" method="post">
	<input type="hidden" name="iC" value="<%=iCurrentpage%>">
	<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
	<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
	<input type="hidden" name="ans_ox" value="<%=sDoc_AnsOX%>">
	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
    <tr height="25" bgcolor="FFFFFF">        
        <td colspan="20" align="center">
         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
		<% else %>[pre]<% end if %>
        <%
			for ix = iStartPage  to iEndPage
				if (ix > iTotalPage) then Exit for
				if Cint(ix) = Cint(iCurrentpage) then
		%>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
		<%		else %>
			<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
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

<%
	set cooperatelist = nothing
%>
<script type="text/javascript">
document.bgColor = "white";
</script>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
