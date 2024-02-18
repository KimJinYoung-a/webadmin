<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전자결재 폼 선택
' History : 2011.03.14 정윤정  생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp"-->
<!-- #include virtual="/lib/classes/approval/araplinkedmsCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->
<%
 Dim iRectMenu
 iRectMenu = "M000"

Dim clsedms, clsALE, arrList, intLoop
Dim icateidx1, icateidx2,sEdmsName
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim iFormType
Dim  sARAPNM,sedmsNM

 	iFormType =  requestCheckvar(Request("rdoT"),1)
 	IF iFormType ="" THEN iFormType = 2

	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

%>
<html>
<head>
<!-- #include virtual="/admin/approval/eapp/eappheader.asp"-->
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.iCP.value=pg;
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


//전자결재 폼선택방법 변경
function jsChangeForm(FormType){
	 document.frm.submit();
}

 //문서 선택
 function jsSetDoc(edmsidx,isPay){
    isPay="False"; //무조건 그냥 진행 품의 후 결제요청서 작성시 수지항목선택 으로 변경 2013/10/21

 	if(isPay =="False"){	//결제요청서 연동되지 않으면 바로 전자결재등록으로 이동
 		location.href= "regeapp.asp?ieidx="+edmsidx;
 	}else{ //결제요청서 연동되는 경우 개정과목 선택 후 이동
 	 	var winApp = window.open("/admin/approval/arap_edms/popGetArapEdms.asp?ieidx="+edmsidx,"popApp","width=600, height=600, resizable=yes");
 		winApp.focus();
 	}
 }

 //폼선택 완료-> 결재페이지 이동
 function jsSelectEApp(iaidx,ieidx){
	location.href= "regeapp.asp?iAidx="+iaidx+"&ieidx="+ieidx;
 }

 //검색
 function jsSearch(){
 	document.frm.iC2.value = $("#selC2").val();  //검색시 ajax 페이지 값 넘겨주기
 	document.frm.submit();
 }
//-->
</script>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" cellpadding="0" cellspacing="0"  border="0">
<tr>
	<td valign="top">
		<table width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<table width="100%"  cellpadding="0" cellspacing="0" class="a" border="0">
				<form name="frm" method="get" action="regeappform.asp">
				<input type="hidden" name="iAidx" value="">
				<input type="hidden" name="ieidx" value="">
				<input type="hidden" name="iC2" value="">
				<input type="hidden" name="iCP" value="1">
				<tr>
					<td>
						<table width="100%"  cellpadding="5"	 cellspacing="1" class="a" border="0" bgcolor="<%= adminColor("tablebg") %>">
						<tr>
							<td width="100"  bgcolor="#DDDDFF" align="center">전자결재 폼선택</td>
							<td  bgcolor="#FFFFFF" >
							   <input type="radio" name="rdoT" value="1" <%IF iFormType = "1" THEN%>checked<%END IF%> onClick="jsChangeForm(1);">전체문서
							   <input type="radio" name="rdoT" value="2" <%IF iFormType = "2" THEN%>checked<%END IF%> onClick="jsChangeForm(2);">결제관련문서
							   <input type="radio" name="rdoT" value="3" <%IF iFormType = "3" THEN%>checked<%END IF%> onClick="jsChangeForm(2);">최근사용문서
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td style="padding-top:10px;">
					<!-- #include virtual="/admin/approval/eapp/include_regeappform_list.asp" -->
					</td>
				</tr>
				</form>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->