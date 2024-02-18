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
dim idx
dim TopTitle
dim QuizDescription
dim BackGroundImage
dim MWbackgroundImage
dim PCWBackGroundImage
dim QuestionHintNumber
dim TotalMileage
dim QuizStartDate
dim QuizEndDate
dim TotalQuestionCount
dim StartDescription
dim AdminRegister
dim AdminName
dim AdminModifyer
dim AdminModifyerName
dim RegistDate
dim LastUpDate
dim QuizStatus
dim productEvtNum
dim endAlertTxt
dim waitingAlertTxt
Dim userid, encUsrId, tmpTx, tmpRn, tenQuizO, totalParticipants, totalWinner, totalQuestions, mileagePerPerson
userid = session("ssBctId")
set tenQuizO = new TenQuiz

'// ���� ����
Dim mode
dim chasu

'// �Ķ���� ����
idx = requestCheckvar(request("idx"),16) 
mode = "mileagediv"

dim tenQuizItem
set tenQuizItem = new TenQuiz
tenQuizItem.FRectIdx = idx
tenQuizItem.GetOneContent()

idx					= tenQuizItem.FoneItem.Fidx											
chasu				= tenQuizItem.FoneItem.Fchasu								
TopTitle			= tenQuizItem.FoneItem.FTopTitle				
QuizDescription		= tenQuizItem.FoneItem.FQuizDescription					
BackGroundImage		= tenQuizItem.FoneItem.FBackGroundImage					
MWbackgroundImage	= tenQuizItem.FoneItem.FMWbackgroundImage
PCWBackGroundImage	= tenQuizItem.FoneItem.FPCWBackGroundImage
QuestionHintNumber	= tenQuizItem.FoneItem.FQuestionHintNumber					
TotalMileage		= tenQuizItem.FoneItem.FTotalMileage								
QuizStartDate		= tenQuizItem.FoneItem.FQuizStartDate					
QuizEndDate			= tenQuizItem.FoneItem.FQuizEndDate						
TotalQuestionCount	= tenQuizItem.FoneItem.FTotalQuestionCount			
StartDescription	= tenQuizItem.FoneItem.FStartDescription
productEvtNum		= tenQuizItem.FoneItem.FProductEvtNum
AdminRegister		= tenQuizItem.FoneItem.FAdminRegister	
AdminName			= tenQuizItem.FoneItem.FAdminName
AdminModifyer		= tenQuizItem.FoneItem.FAdminModifyer		
AdminModifyerName	= tenQuizItem.FoneItem.FAdminModifyerName	
RegistDate			= tenQuizItem.FoneItem.FRegistDate	
LastUpDate			= tenQuizItem.FoneItem.FLastUpDate	
QuizStatus			= tenQuizItem.FoneItem.FQuizStatus	
endAlertTxt			= tenQuizItem.FoneItem.FEndAlertTxt
waitingAlertTxt		= tenQuizItem.FoneItem.FWaitingAlertTxt

totalParticipants = tenQuizItem.GetNumberOfParticipants(chasu)
totalWinner = tenQuizItem.GetNumberOfWinner(chasu, TotalQuestionCount)

IF totalWinner=0 then
	response.write "<script>alert('������ ����� �����ϴ�.'); window.close();</script>"
else
	mileagePerPerson = formatnumber(TotalMileage/totalWinner, 0)
end if

if mileagePerPerson > 5000 then '��÷�ݾ��� �ʹ� �������
	response.write "<script>alert('�����ڿ��� ���� �ٶ��ϴ�.'); window.close();</script>"
end if
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
function divMileage(){
	if(confirm(`������ �Ѹ� �� <%=mileagePerPerson%>���Դϴ�. �й� �Ͻðڽ��ϱ�?`)){
		var frm = document.frm;		
		frm.action = "tenquizaction.asp"
		frm.submit();
	}	
}
</script>

<form name="frm" method="post">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="chasu" value="<%=chasu%>">
<input type="hidden" name="mode" value="<%=mode%>">

<div class="popWinV17">
	<h2 class="tMar20 subType" style="margin-left:30px;">���ϸ��� �й�</h2>
	<div class="popContainerV17 pad30">
		<table class="tbType1 writeTb tMar10">
			<colgroup>
				<col width="18%" /><col width="" />
			</colgroup>
			<tbody>				
			<tr>
				<th><div>���� <strong class="cRd1"></strong></div></th>
				<td>
					<p><b><%=chasu%></b></p>
				</td>
			</tr>
			<tr>
				<th><div>���ϸ��� <strong class="cRd1"></strong></div></th>
				<td>
					<p><b><%=formatnumber(TotalMileage, 0)%>��</b></p>
				</td>
			</tr>
			<tr>
				<th><div>������/������ <strong class="cRd1"></strong></div></th>
				<td>
					<p><b><span style="color:red"><%=formatnumber(totalWinner,0)%></span>/<%=formatnumber(totalParticipants,0)%></b></p>
				</td>
			</tr>
			<tr>
				<th><div>�δ� ���ϸ���<strong class="cRd1"></strong></div></th>
				<td>
					<p><b><%if totalWinner <> 0 then response.write formatnumber(TotalMileage/totalWinner, 0) else response.write 0 %> ��</b></p>
				</td>
			</tr>								
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<!-- input type="button" value="�̸�����" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
		<input type="button" value="���" onclick="divMileage();" class="cRd1" style="width:100px; height:30px;" />
		<!-- input type="button" value="����" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>	
</div>	
</form>	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->