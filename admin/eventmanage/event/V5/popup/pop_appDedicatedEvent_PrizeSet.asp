<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_appDedicatedEvent_PrizeSet.asp
' Discription : ������ �̺�Ʈ ��÷�� ���� ��� â
' History : 2023.02.07 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/appDedicatedEventCls.asp"-->
<%
dim mode, oAppDedicated, arrList, intLoop, prizeArr
dim evt_code : evt_code = request("evt_code")
dim episode : episode = request("episode")
dim itemid : itemid = request("itemid")

set oAppDedicated = new AppEventCls
oAppDedicated.FRectEventCode = evt_code
oAppDedicated.FRectEpisode = episode
arrList = oAppDedicated.fnGetAppDedicatedPrizeList
if isArray(arrList) then 
    for intLoop = 0 to UBound(arrList,2)
        if intLoop = 0 then
            prizeArr = arrList(0,intLoop)
        else
            prizeArr = prizeArr & "," & arrList(0,intLoop)
        end if
    next
end if
set oAppDedicated = nothing
%>
<script>
function jsEvtSubmit(frm){
    if(frm.prizearr.value==""){
        alert("��÷�� ���̵� ������ּ���.");
        return false;
    }
    frm.action="appDedicatedItem_process.asp";
	frm.submit();
}
</script>
<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="prize">
<input type="hidden" name="episode" value="<%=episode%>">
<input type="hidden" name="itemid" value="<%=itemid%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>������ ������ ��÷�� ����</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>��÷�� ���̵� ���</th>
                    <td>
                        <% if prizeArr = "" then %>
                        <textarea name="prizearr" rows="8" cols="50" placeholder="��÷�� ���̵� ��ǥ�� �����Ͽ� ������ּ���."></textarea>
                        <% else %>
                        <%=prizeArr%>
                        <% end if %>
                    </td>
                </tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">���</button>
        <% if prizeArr = "" then %>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">����</button>
        <% end if %>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->