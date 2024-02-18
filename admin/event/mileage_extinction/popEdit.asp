<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/EventMileageCls.asp" -->
<%
Dim userid, id, exCls, modTxt, artMsg

dim jukyo
dim jukyocd
dim startdate
dim enddate
dim chkDays
dim useyn

userid = session("ssBctId")
id = request("id")

dim mode
dim lastDays
 mode = chkIIF(id <> "", "mod", "add")

if mode = "mod" Then
	set exCls = new MileageExtinctionCls
	exCls.FRectSubIdx = id
	exCls.GetOneSubItem()

	jukyo = exCls.FOneItem.task_jukyo
	jukyocd = exCls.FOneItem.task_jukyocd
	startdate = exCls.FOneItem.task_startdate
	enddate = exCls.FOneItem.task_enddate
	chkDays = exCls.FOneItem.task_chkDays
	useyn	 = exCls.FOneItem.task_useyn
	
	lastDays = datediff("d", date(),dateadd("d", chkDays + 1, enddate))
	dim tmpUseYn : tmpUseYn =tmpUseYn
	if lastDays <= chkDays then
		modTxt = "style=""background-color:#eeeded"" readonly disabled"
		artMsg = "진행중인 소멸 작업은 사용 여부를 제외한 다른 정보 수정이 불가합니다."
	end if	
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
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<link rel="stylesheet" href="/resources/demos/style.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript">
$(function(){
    $("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		minDate: "<%=dateserial(year(now),month(now)-6,1)%>"
    });
    $("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		maxDate: "<%=dateadd("d",-1,dateserial(year(now),month(now)+7,1))%>"
    });
})
function validate(){
	var chkRes = true
	$(".form-table input").each(function(idx, el){
		if(el.value == ''){
			alert('필수 사항을 넣어주세요.');
			el.focus();
			chkRes = false
			return false;
		}
		if(el.name == 'jukyocd'){
			var reservedCodes = [300000, 400000, 400001, 600000, 1, 2, 99999, 999, 1100, 1000]
			chkRes = !reservedCodes.some((item, idx, arr) => {
				return item == el.value;
			});
			if(!chkRes){
				alert('예약된 적요코드입니다. 다른 코드를 넣어주세요.')
				return false;
			}
		}
		if(el.name == 'chkDays'){			
			if(el.value < 5){
				alert('최소 체크기간은 5일입니다.')
				el.value = 5
				chkRes = false
				return false;
			}
		}		
	})
	return chkRes
}
function submitContent(){
	if(!validate()) return false;
	console.log('hi')

	var frm = document.frm
	frm.action = "extinction_act.asp"	
	frm.submit()
}
</script>

<form name="frm" method="post">
<input type="hidden" name="adminid" value="<%=session("ssBctId")%>">
<input type="hidden" name="mode" value="<%=mode%>" />
<input type="hidden" name="id" value="<%=id%>" />
<div class="popWinV17">
<% If mode = "add" Then %>
	<h2 class="tMar20 subType" style="margin-left:30px;">작업 추가</h2>
<% Else %>
	<h2 class="tMar20 subType" style="margin-left:30px;">작업 수정</h2>
<% End if %>
	<div class="popContainerV17 pad30">
		<span class="cOr1"><%=artMsg%></span>
		<table class="tbType1 writeTb tMar10 form-table">
			<colgroup>
				<col width="20%" /><col width="" />
			</colgroup>
			<tbody>
				<tr>
					<th><div>소멸 적요<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" <%=modTxt%> name="jukyo" class="formTxt" style="width:50%;" value="<%=jukyo%>"/></p>
						<p class="tPad05 fs11 cGy1">- 프론트에 보여지는 소멸 로그입니다. ex) 3333마일리지 이벤트 소멸</p>
					</td>
				</tr>
				<tr>
					<th><div>이벤트코드(적요코드)<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" <%=modTxt%> name="jukyocd" class="formTxt" style="width:15%;" value="<%=jukyocd%>"/></p>
						<p class="tPad05 fs11 cGy1">- 마일리지 이벤트 코드입니다.</p>
					</td>
				</tr>
				<tr>
					<th><div>이벤트 진행 기간<strong class="cRd1">*</strong></div></th>
					<td>
						시작일: <input type="text" <%=modTxt%> id="sDt"  name="startdate" class="formTxt" style="width:15%;" value="<%=startdate%>" readonly/>
						<br> 종료일: <input type="text" <%=modTxt%> id="eDt"  name="enddate" class="formTxt" style="width:15%;" value="<%=enddate%>" readonly/>
						<p class="tPad05 fs11 cGy1">- 이벤트 진행 기간입니다. 소멸작업은 종료일 다음날부터 시작됩니다.<br />- 종료일은 현재부터 최대 6개월입니다.</p>
					</td>
				</tr>
				<tr>
					<th><div>마일리지 체크 기간<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" <%=modTxt%> name="chkDays" min="1" max="10" class="formTxt" style="width:15%;" value="<%=chkDays%>"/>일</p>
						<p class="tPad05 fs11 cGy1">- 이벤트 시작일 ~ 종료일까지 마일리지를 사용한 구매이력이 있을 시 소멸 대상에서 제외됩니다. 단, 체크 기간동안 구매 취소가 일어났을 경우 다시 소멸대상으로 간주되어 받은 마일리지는 소멸됩니다. 체크 기간은 무통장 입금 결제방식을 고려하여 기본 최소 5일이며 설정 가능합니다.</p>
					</td>
				</tr>
				<tr>
					<th><div>사용여부<strong class="cRd1">*</strong></div></th>
					<td>
							<input type="radio" <%=chkIIF(lastDays < 0, modTxt, "")%> name="useyn" value=1 <%=chkIIF(useyn="1" or useyn="", "checked", "")%>> 사용<br>
							<input type="radio" <%=chkIIF(lastDays < 0, modTxt, "")%> name="useyn" value=0 <%=chkIIF(useyn="0", "checked", "")%>> 사용안함
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="취소" onclick="window.close();" style="width:100px; height:30px;" />
		<input type="button" value="저장" onclick="submitContent();" class="cRd1" style="width:100px; height:30px;" />
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->