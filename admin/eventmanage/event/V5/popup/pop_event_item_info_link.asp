<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : pop_event_item_info_link.asp
' Discription : I��(������) �̺�Ʈ ������ ��ǰ ���� ���
' History : 2022.06.16 ������
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
Dim cEvtCont, itemArray, eventType
Dim eventCode, itemArrayInfo

eventCode = Request("evt_code")

IF eventCode <> "" THEN
	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eventCode	'�̺�Ʈ �ڵ�
	'�̺�Ʈ ���� ��������
	itemArrayInfo=cEvtCont.fnGetItemInfoLinkEvent
    If isArray(itemArrayInfo) Then
        eventType = itemArrayInfo(0,0)
		itemArray = itemArrayInfo(1,0)
    end if
	set cEvtCont = nothing
end if
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";

function jsEvtSubmit(frm){
    //ä�μ��� ���� Ȯ��
    if (!frm.eventType.value){
        alert("�̺�Ʈ Ÿ���� �������ּ���.");
        frm.eventType.focus();
        return false;
    }

    if(!frm.itemArray.value){
        alert("������ ������ �Է����ּ���");
        frm.itemArray.focus();
        return false;
    }

    frm.action="iteminfolink_process.asp";
    frm.submit();
}

</script>

<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="eventCode" value="<%=eventCode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>��ǰ ���� ��ȹ�� ����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>�̺�Ʈ Ÿ��</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="1" <% if eventType=1  then %> checked<% end if %>>
								�����Shop
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="2" <% if eventType=2  then %> checked<% end if %>>
								��¦Ư��
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="3" <% if eventType=3  then %> checked<% end if %>>
								�÷�������
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr>
					<th>��ǰ�ڵ� ����</th>
					<td>
						<div class="formInline">
							<input type="text" name="itemArray" class="formControl formControl550" placeholder="��ǰ�ڵ�" maxlength="256" value="<%=itemArray%>" style="width:600px">
						</div>
						<p class="tMar15 cPk2 fs12">���� ���� �޸��� �������ּ���.(ex : 4710529,4710517,4710463)</p>
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->