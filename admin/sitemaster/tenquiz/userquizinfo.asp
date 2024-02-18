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
'// ���� ����
dim i, chasu, userid, questionNumber, answer, userAnswer, result

'// �Ķ���� ����
chasu = request("chasu")
userid = request("userid")

dim tenQuizUserQuizList
set tenQuizUserQuizList = new TenQuiz
tenQuizUserQuizList.FRectUserId = userid
tenQuizUserQuizList.FRectChasu = chasu
tenQuizUserQuizList.GetUserAnswerList()
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
	if(confirm(`����: ${chasu}, ����ھ��̵�: ${userid}�Դϴ�. ���� �̷��� �����Ͻðڽ��ϱ�?`)){
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
	<h2 class="tMar20 subType" style="margin-left:30px;">����� ��������</h2>
	<div class="popContainerV17 pad30">
		<table class="tbType1 writeTb tMar10">
			<colgroup>
				<col width="18%" /><col width="" />
			</colgroup>
			<tbody>				
				<tr>
					<th><div>���̵�</div></th>
					<td colspan=2>
						<p><b><%=userid%></b></p>
					</td>					
					<td>
						<p><b><button type="button" onclick="deleteUserData();">�������������</button></b></p>
					</td>					
				</tr>							
				<tr>
					<th><div>����</div></th>
					<td colspan=3>
						<p><b><%=chasu%></b></p>
					</td>
				</tr>				
				<tr>
					<th><div>���</div></th>
					<td colspan=3>
						<p><b><%=tenQuizUserQuizList.FItemList(0).FAuserscore%>/<%=tenQuizUserQuizList.FItemList(0).FAtotalquestioncount%></b></p>
					</td>
				</tr>								
				<tr>
					<th><div> ������ȣ</div></th>
					<th><div>������ ��</div></th>					
					<th><div>����</div></th>					
					<th><div>���</div></th>					
				</tr>								
			<%  dim tmpUserAnswer
				for i=0 to tenQuizUserQuizList.FTotalCount-1  
					if tenQuizUserQuizList.FItemList(i).FAuserAnswer = -1 then
						tmpUserAnswer = "�ð��ʰ�"
					elseif tenQuizUserQuizList.FItemList(i).FAuserAnswer = 0 then
						tmpUserAnswer = "��Ʈ��ũ����"
					else
						tmpUserAnswer = tenQuizUserQuizList.FItemList(i).FAuserAnswer
					end if
			%>
				<tr style="text-align:center">
					<th><div><%=tenQuizUserQuizList.FItemList(i).FAquestionNumber%>������ <strong class="cRd1"></strong></div></th>
					<td>
						<p><b><%=tmpUserAnswer%></b></p>
					</td>
					<td>
						<p><b><%=tenQuizUserQuizList.FItemList(i).FAanswer%></b></p>
					</td>
					<td>
						<p><b><%=chkIIF(tenQuizUserQuizList.FItemList(i).FAresult, "<span style=""color:blue"">O</span>", "<span style=""color:red"">X</span>") %></b></p>
					</td>					
				</tr>					
			<% next %>						
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<!-- input type="button" value="�̸�����" onclick="" class="cBl2" style="width:100px; height:30px;" / -->
		<input type="button" value="Ȯ��" onclick="window.close();" style="width:100px; height:30px;" />		
		<!-- input type="button" value="����" onclick="" class="cRd1" style="width:100px; height:30px;" / -->
	</div>	
</div>	
</form>	
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->