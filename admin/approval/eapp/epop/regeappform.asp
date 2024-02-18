<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 폼 선택
' History : 2011.03.14 정윤정  생성
'						2013/10/21 무조건 그냥 진행 품의 후 결제요청서 작성시 수지항목선택 으로 변경 2013/10/21
'						2013.12.3  정윤정 변경 - 폼선택없이 검색조건으로  문서선택  처리 
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp"--> 
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim clsedms, clsALE, arrList, intLoop
Dim icateidx1, icateidx2,sARAPNM,sedmsNM,suserid
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage 
Dim iStartPage,iEndPage,iX,iPerCnt
		iPageSize = 20
		iCurrPage = requestCheckvar(Request("iCP"),10)
		if iCurrPage="" then iCurrPage=1
		
	  icateidx1 = requestCheckvar(Request("selC1"),10)
		icateidx2 = requestCheckvar(Request("iC2"),10)
		sedmsNM = requestCheckvar(Request("sENM"),128)
		sARAPNM = requestCheckvar(Request("sANM"),60)
		suserid = requestCheckvar(Request("sUID"),32)
		if icateidx1 = "" then icateidx1 = 0
		if icateidx2 = "" then icateidx2= 0

		Set clsALE = new CArapLinkEdms
			clsALE.Fcateidx1 = icateidx1
			clsALE.Fcateidx2	= icateidx2 
			clsALE.FEdmsName 	= sedmsNM
			clsALE.FARAP_NM 	= sARAPNM
			clsALE.FCurrPage	= iCurrPage
			clsALE.FPageSize	= iPageSize
			clsALE.FadminId		= suserid
			arrList = clsALE.fnGetPartTimeEappArapLinkEdmsList
			iTotCnt = clsALE.FTotCnt
			iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript"> 
// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.iCP.value=pg;
		document.frm.submit();
	}

 //검색
 function jsSearch(){ 
 	document.frm.iC2.value = $("#selC2").val();  //검색시 ajax 페이지 값 넘겨주기
 	document.frm.submit();
 }
 
//카테고리 선택
$(document).ready(function(){
	$("#selC1").change(function(){
		var iValue = $("#selC1").val();
		var url="/admin/approval/edms/ajaxCategory.asp";
		var params = "sMode=CL&ipcidx="+iValue ;

		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("#divCL").html(args);
		 	},
		 	error:function(e){
		 		alert("데이터로딩에 문제가 생겼습니다. 시스템팀에 문의해주세요");
		 		//alert(e.responseText);
		 	}
		 });
	});
});

 //폼선택 완료-> 결재페이지 이동
 function jsSelectEApp(iaidx,ieidx,inum){   
	location.href= "regeapp.asp?iAidx="+iaidx+"&ieidx="+ieidx;
 }


</script>
<style>
	FORM {display:inline;}
	</style>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr>
	<td valign="top"> 
		<table width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td height="30" valign="bottom" >[ 전자결재 폼선택 ]</td>
		</tr>	
		<tr>
			<td>
			<form name="frm" method="get" action="regeappform.asp">
			<input type="hidden" name="iAidx" value="">
			<input type="hidden" name="ieidx" value="">
			<input type="hidden" name="iC2" value="">
			<input type="hidden" name="iCP" value="1">
			<input type="hidden" name="sUID" value="<%=suserid%>">
				<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr align="center" bgcolor="#FFFFFF" >
					<td width="100" height="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">검색 조건</td>
					<td align="left">
						<%	Set clsedms = new Cedms%> 
						문서카테고리:
						<select name="selC1" id="selC1">
						<option value="0">전체</option>
						<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
						</select> 
						>
						<span id="divCL">
						<select name="selC2" id="selC2">
						<option value="0">전체</option>
					<% 	IF icateidx1 > 0 THEN	'대카테고리 선택 후 중카테고리 선택가능하게
							clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2
						END IF
					%>
						</select>
						</span>
						<% Set clsedms = nothing %>  
						&nbsp;&nbsp;문서명 : <input type="text" name="sENM" value="<%=sEdmsNM%>" size="20">
						&nbsp;&nbsp;
					</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="jsSearch();"></td>
				</tr>
				<Tr bgcolor="#FFFFFF" >
					<td>
						수지항목: <input type="text" name="sANM" value="<%=sARAPNM%>" size="20">   
					</td>
				</tr>
				</table> 
			</form>	
			</td>
		</tr>
		<tr>
		<td><br>검색결과 : <b><%=iTotCnt%></b> &nbsp;
			<!-- 상단 띠 시작 -->
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				<!-- <td>idx</td> -->
				<td>문서코드</td>
				<td>문서카테고리</td>
				<td>문서명</td>
				<td>최종결재자</td>
				<td>결제요청서</td>
				<td>수지항목</font></td>
				<td>연결계정과목</font></td>
			</tr>
			<%
				IF isArray(arrList) THEN
					For intLoop = 0 To UBound(arrList,2) 
			%>
					<tr height=30 align="center" bgcolor="#FFFFFF"> 
						<td nowrap><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><%=arrList(1,intLoop)%></td>
						<td align="left"><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><%=arrList(4,intLoop)%> > <%=arrList(5,intLoop)%></td>
						<td><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><%=arrList(6,intLoop)%></td>
						<td><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><%=arrList(8,intLoop)%></td>
						<td><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><%IF arrList(9,intLoop) THEN%>Y<%ELSE%>N<%END IF%></td>
						<td align="left"><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><% if isNULL(arrList(11,intLoop)) then%><% if  arrList(9,intLoop) then %><font color="gray">결제요청시 선택</font><%end if%><% else %><font color="blue">[<%=arrList(11,intLoop)%>] <%=arrList(12,intLoop)%></font><% end if %></a></td>
						<td align="left"><a href="javascript:jsSelectEApp('<%=arrList(11,intLoop)%>','<%=arrList(0,intLoop)%>','<%=intLoop%>');"><% if isNULL(arrList(11,intLoop)) then%><% if  arrList(9,intLoop) then %><font color="gray">결제요청시 선택</font><%end if%><% else %><font color="blue">[<%=arrList(14,intLoop)%>] <%=arrList(13,intLoop)%></font><% end if %></a></td>
					</tr>
			<%
					Next
				ELSE
			%>
					<tr height=5 align="center" bgcolor="#FFFFFF"><td colspan="10">등록된 내용이 없습니다.</td></tr>
			<% END IF %>
			</table>
		</td>
	</tr><!-- 페이지 시작 -->
	<!-- #include virtual="/admin/approval/eapp/include_regeappform_list_paging.asp" --> 
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->