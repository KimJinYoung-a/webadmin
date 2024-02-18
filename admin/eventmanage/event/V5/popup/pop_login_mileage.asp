<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : pop_login_mileage.asp
' Discription : I��(������) �̺�Ʈ ������ �α��� ���ϸ��� ���
' History : 2021.11.26 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
Dim cEvtCont, mileagePoint, jukyo
Dim eCode, emdid, emdnm, mileageInfo

eCode = Request("evt_code")

if emdid = "" then 
    emdid = session("ssBctId")
    emdnm = session("ssBctCname")
end if

IF eCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	mileageInfo=cEvtCont.fnGetLoginMileageEvent
    If isArray(mileageInfo) Then
        mileagePoint = mileageInfo(0,0)
        jukyo = mileageInfo(1,0)
    end if
	set cEvtCont = nothing
end if
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";

function jsEvtSubmit(frm){
    //ä�μ��� ���� Ȯ��
    if (!frm.mileagePoint.value){
        alert("���ϸ��� ����Ʈ�� �Է����ּ���.");
        frm.mileagePoint.focus();
        return false;
    }

    if(!frm.jukyo.value){
        alert("���並 �Է����ּ���");
        frm.jukyo.focus();
        return false;
    }

    frm.action="loginmileage_process.asp";
    frm.submit();
}

</script>

<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=eCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>�α��� ���ϸ��� ��ȹ�� ����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr id="mileagediv">
					<th>���ϸ��� �̺�Ʈ ���� ����</th>
					<td>
						<div class="formInline">
							<input type="text" name="mileagePoint" class="formControl formControl550" placeholder="���� ����Ʈ" maxlength="4" value="<%=mileagePoint%>">
						</div><br>
						<div class="formInline">
							<input type="text" name="jukyo" class="formControl formControl550" placeholder="����" maxlength="128" value="<%=jukyo%>" style="width:600px">
						</div>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
	</div>
</div>
</form>
<script>
<% if eCode ="" then %>
$(function() {
	$("select[name='eventlevel']").val("3").attr("selected","selected");
});
<% end if %>
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->