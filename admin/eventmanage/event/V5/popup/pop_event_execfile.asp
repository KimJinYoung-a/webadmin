<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_event_execfile.asp
' Discription : I��(������) �̺�Ʈ ��ȹ�� ��Ƽ������ �������� ����
' History : 2019.10.14 ������
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
Dim cEvtCont, winmode, idx
Dim eCode, ExecFile, menuidx, isusing

eCode = Request("eC")
menuidx = Request("menuidx")
winmode = Request("wm")
if winmode="" then winmode="M"
IF eCode <> "" THEN
    set cEvtCont = new ClsMultiContentsMenu
    cEvtCont.FRectEvtCode = eCode
    cEvtCont.FRectIDX = menuidx	'��Ƽ ������ �޴� idx
    cEvtCont.FRectDevice = winmode
    cEvtCont.fnGetMultiContentsImgText
    idx = cEvtCont.Fidx
    ExecFile = cEvtCont.FImgURL
    isusing = cEvtCont.Fisusing
    set cEvtCont = nothing
else

end if
%>
<script type="text/javascript" > 
window.document.domain = "10x10.co.kr";
function jsEvtSubmit(frm){
    if(!frm.sEFP.value){
        alert("�������� ��θ� �Է����ּ���.");
        return;
    }
    else{
	    frm.submit();
    }
}
</script>
<form name="frmEvt" method="post" style="margin:0px;" action="execfile_process.asp">
<% if idx<>"" then %>
<input type="hidden" name="imod" value="TU">
<% else %>
<input type="hidden" name="imod" value="TI">
<% end if %>
<input type="hidden" name="evt_code" value="<%=eCode%>">
<input type="hidden" name="device" value="<%=winmode%>">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menuidx" value="<%=menuidx%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>��������</h1>
	</div>
	<div class="popContV19">
		<div class="tabV19">
			<ul>
				<li class="<% if winmode="M" then %>selected<% end if %>"><a href="pop_event_execfile.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=M">Mobile / App</a></li>
				<li class="<% if winmode="W" then %>selected<% end if %>"><a href="pop_event_execfile.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>&wm=W">PC</a></li>
			</ul>
		</div>
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <th>Exec File (��������)</th>
                    <td>
                        <input type="text" class="formControl" placeholder="��������" name="sEFP" value="<%=ExecFile%>">
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="isusing" value="N" <%if isusing="N" then%>checked<%end if%>>
                                �����
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="isusing" value="Y" <%if isusing="Y" or isusing="" then%>checked<%end if%>>
                                ����
                                <i class="inputHelper"></i>
                            </label>
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
<form method="post" name="ibfrm">
	<input type="hidden" name="idx">
	<input type="hidden" name="mode">
</form>
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->