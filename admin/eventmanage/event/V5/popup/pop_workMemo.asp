<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_workMemo.asp
' Discription : I��(������) �̺�Ʈ �۾� ���� ���� ��
' History : 2019.01.24 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<%
dim cEvtCont, fullefwd, efwd
dim eCode : eCode = Request("eC")
set cEvtCont = new ClsEvent
cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
cEvtCont.fnGetEventDisplay
fullefwd = nl2br(db2html(cEvtCont.FEFwd))
efwd = db2html(cEvtCont.FEFwd)
set cEvtCont = nothing
%>
<script>
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    frm.action="workmemo_process.asp";
    frm.submit();
}
function fnChangeEdit(){
	$("#readcon").hide();
	$("#writecon").show();
	$("#btnhide").show();
}
</script>
<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="imod" value="WU">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>�۾����޻���</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>�۾����޻���<br>(���� Ŭ���� ���� ����)</th>
					<td onClick="fnChangeEdit();">
						<span id="readcon">
						<%=fullefwd%>
						</span>
						<span id="writecon" style="display:none">
						<textarea name="tFwd" rows="16" cols="50" placeholder="�۾��� ���޻���"><%=efwd%></textarea>
						</span>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19" style="display:none" id="btnhide">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->