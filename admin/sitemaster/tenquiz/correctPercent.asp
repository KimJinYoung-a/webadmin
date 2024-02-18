<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/sitemasterclass/TenQuizCls.asp" -->
<%
'// 변수 선언
dim i, chasu, userid, questionNumber, answer, userAnswer, result
dim itemArr

'// 파라메터 접수
chasu = request("chasu")
userid = request("userid")

dim tenQuizList
set tenQuizList = new TenQuiz
itemArr = tenQuizList.getQuizCorrectPercent(chasu)
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script type="text/javascript">
function deleteUserData(){	
	var frm = document.frm;
	var userid = frm.deleteUserId.value;
	var chasu = frm.chasu.value;
	frm.action = "tenquizaction.asp";
	if(confirm(`차수: ${chasu}, 사용자아이디: ${userid}입니다. 퀴즈 이력을 삭제하시겠습니까?`)){
		frm.submit();
	}
	return false;
}
</script>

<form name="frm" method="post">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="deleteUserData">
<input type="hidden" name="deleteUserId" value="<%=userid%>">
<input type="hidden" name="chasu" value="<%=chasu%>">
<div class="popWinV17">
	<h2 class="tMar20 subType" style="margin-left:30px;">사용자 퀴즈정보</h2>
	<div class="popContainerV17 pad30">
		<table class="tbType1 writeTb tMar10">
			<colgroup>
				<col width="18%" /><col width="" />
			</colgroup>
			<tbody>				
				<tr>
					<th><div>차수</div></th>
					<td colspan=2>
						<p><b><%=chasu%></b></p>
					</td>
				</tr>								
				<tr>
					<th><div> 문제번호</div></th>					
					<th><div>정답</div></th>					
					<th><div>정답률</div></th>					
				</tr>								
			<%  
				for i=0 to ubound(itemArr, 2)
			%>
				<tr style="text-align:center">
					<th><div><%=itemArr(0, i)%>번문제 <strong class="cRd1"></strong></div></th>					
					<td>
						<p><b><%=itemArr(1, i)%></b></p>
					</td>
					<td>
						<p><b style="color:red"><%=itemArr(2, i)%> %</b></p>
					</td>					
				</tr>					
			<% next %>						
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<!-- input type="button" value="미리보기" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="확인" onclick="window.close();" style="width:100px; height:30px;" />		
		<!-- input type="button" value="수정" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>	
</div>	
</form>	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->