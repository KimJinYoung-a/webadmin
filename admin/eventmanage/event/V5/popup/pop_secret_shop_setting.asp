<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_secret_shop_setting.asp
' Discription : 비밀의 샵 상품 설정 등록 창
' History : 2023.05.09 정태훈
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
dim mode, oAppEvent, itemarr
dim evt_code : evt_code = request("evt_code")

set oAppEvent = new AppEventCls
    oAppEvent.FRectEventCode = evt_code
    itemarr = oAppEvent.fnGetSecretShopItemInfo
set oAppEvent = nothing

if itemarr(0,0) <> "" then 
    mode = "Modify"
else
    mode = "Add"
end if


%>
<script type="text/javascript" src="/js/jquery.form.min.js"></script> 
<script>
function jsEvtSubmit(frm){
    if(frm.itemarr.value==""){
        alert("상품 번호를 등록해주세요.");
        return false;
    }
    frm.action="secretShop_process.asp";
	frm.submit();
}
</script>
<form name="frmEvt" method="post" style="margin:0px;">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="mode" value="<%=mode%>">
<div class="popV19">
	<div class="popHeadV19">
		<h1>비밀의 샵 상품 설정</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>상품등록</th>
					<td>
                        <input type="text" class="formControl formControl550" placeholder="상품 번호를 등록해주세요.(ex:1122,1214 콤마구분)" name="itemarr" id="itemarr" value="<%=itemarr(0,0)%>">
          			</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsEvtSubmit(this.form);">저장</button>
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->