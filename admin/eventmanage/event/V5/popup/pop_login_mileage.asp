<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : pop_login_mileage.asp
' Discription : I형(통합형) 이벤트 마케팅 로그인 마일리지 등록
' History : 2021.11.26 정태훈
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
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
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
    //채널선택 여부 확인
    if (!frm.mileagePoint.value){
        alert("마일리지 포인트를 입력해주세요.");
        frm.mileagePoint.focus();
        return false;
    }

    if(!frm.jukyo.value){
        alert("적요를 입력해주세요");
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
		<h1>로그인 마일리지 기획전 정보</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr id="mileagediv">
					<th>마일리지 이벤트 설정 유무</th>
					<td>
						<div class="formInline">
							<input type="text" name="mileagePoint" class="formControl formControl550" placeholder="지급 포인트" maxlength="4" value="<%=mileagePoint%>">
						</div><br>
						<div class="formInline">
							<input type="text" name="jukyo" class="formControl formControl550" placeholder="적요" maxlength="128" value="<%=jukyo%>" style="width:600px">
						</div>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);return false;">저장</button>
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