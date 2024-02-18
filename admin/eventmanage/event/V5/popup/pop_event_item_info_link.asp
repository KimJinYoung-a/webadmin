<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : pop_event_item_info_link.asp
' Discription : I형(통합형) 이벤트 마케팅 상품 연동 등록
' History : 2022.06.16 정태훈
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
	cEvtCont.FECode = eventCode	'이벤트 코드
	'이벤트 내용 가져오기
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
    //채널선택 여부 확인
    if (!frm.eventType.value){
        alert("이벤트 타입을 선택해주세요.");
        frm.eventType.focus();
        return false;
    }

    if(!frm.itemArray.value){
        alert("아이템 정보를 입력해주세요");
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
		<h1>상품 연동 기획전 정보</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>이벤트 타입</th>
					<td>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="1" <% if eventType=1  then %> checked<% end if %>>
								비밀의Shop
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="2" <% if eventType=2  then %> checked<% end if %>>
								깜짝특가
								<i class="inputHelper"></i>
							</label>
						</div>
						<div class="formInline">
							<label class="formCheckLabel">
								<input type="radio" class="formCheckInput" name="eventType" id="radio7b" value="3" <% if eventType=3  then %> checked<% end if %>>
								플러스세일
								<i class="inputHelper"></i>
							</label>
						</div>
					</td>
				</tr>
				<tr>
					<th>상품코드 정보</th>
					<td>
						<div class="formInline">
							<input type="text" name="itemArray" class="formControl formControl550" placeholder="상품코드" maxlength="256" value="<%=itemArray%>" style="width:600px">
						</div>
						<p class="tMar15 cPk2 fs12">공백 없이 콤마로 구분해주세요.(ex : 4710529,4710517,4710463)</p>
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->