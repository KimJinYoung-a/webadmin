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
		artMsg = "�������� �Ҹ� �۾��� ��� ���θ� ������ �ٸ� ���� ������ �Ұ��մϴ�."
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
			alert('�ʼ� ������ �־��ּ���.');
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
				alert('����� �����ڵ��Դϴ�. �ٸ� �ڵ带 �־��ּ���.')
				return false;
			}
		}
		if(el.name == 'chkDays'){			
			if(el.value < 5){
				alert('�ּ� üũ�Ⱓ�� 5���Դϴ�.')
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
	<h2 class="tMar20 subType" style="margin-left:30px;">�۾� �߰�</h2>
<% Else %>
	<h2 class="tMar20 subType" style="margin-left:30px;">�۾� ����</h2>
<% End if %>
	<div class="popContainerV17 pad30">
		<span class="cOr1"><%=artMsg%></span>
		<table class="tbType1 writeTb tMar10 form-table">
			<colgroup>
				<col width="20%" /><col width="" />
			</colgroup>
			<tbody>
				<tr>
					<th><div>�Ҹ� ����<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="text" <%=modTxt%> name="jukyo" class="formTxt" style="width:50%;" value="<%=jukyo%>"/></p>
						<p class="tPad05 fs11 cGy1">- ����Ʈ�� �������� �Ҹ� �α��Դϴ�. ex) 3333���ϸ��� �̺�Ʈ �Ҹ�</p>
					</td>
				</tr>
				<tr>
					<th><div>�̺�Ʈ�ڵ�(�����ڵ�)<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" <%=modTxt%> name="jukyocd" class="formTxt" style="width:15%;" value="<%=jukyocd%>"/></p>
						<p class="tPad05 fs11 cGy1">- ���ϸ��� �̺�Ʈ �ڵ��Դϴ�.</p>
					</td>
				</tr>
				<tr>
					<th><div>�̺�Ʈ ���� �Ⱓ<strong class="cRd1">*</strong></div></th>
					<td>
						������: <input type="text" <%=modTxt%> id="sDt"  name="startdate" class="formTxt" style="width:15%;" value="<%=startdate%>" readonly/>
						<br> ������: <input type="text" <%=modTxt%> id="eDt"  name="enddate" class="formTxt" style="width:15%;" value="<%=enddate%>" readonly/>
						<p class="tPad05 fs11 cGy1">- �̺�Ʈ ���� �Ⱓ�Դϴ�. �Ҹ��۾��� ������ ���������� ���۵˴ϴ�.<br />- �������� ������� �ִ� 6�����Դϴ�.</p>
					</td>
				</tr>
				<tr>
					<th><div>���ϸ��� üũ �Ⱓ<strong class="cRd1">*</strong></div></th>
					<td>
						<p><input type="number" <%=modTxt%> name="chkDays" min="1" max="10" class="formTxt" style="width:15%;" value="<%=chkDays%>"/>��</p>
						<p class="tPad05 fs11 cGy1">- �̺�Ʈ ������ ~ �����ϱ��� ���ϸ����� ����� �����̷��� ���� �� �Ҹ� ��󿡼� ���ܵ˴ϴ�. ��, üũ �Ⱓ���� ���� ��Ұ� �Ͼ�� ��� �ٽ� �Ҹ������� ���ֵǾ� ���� ���ϸ����� �Ҹ�˴ϴ�. üũ �Ⱓ�� ������ �Ա� ��������� ����Ͽ� �⺻ �ּ� 5���̸� ���� �����մϴ�.</p>
					</td>
				</tr>
				<tr>
					<th><div>��뿩��<strong class="cRd1">*</strong></div></th>
					<td>
							<input type="radio" <%=chkIIF(lastDays < 0, modTxt, "")%> name="useyn" value=1 <%=chkIIF(useyn="1" or useyn="", "checked", "")%>> ���<br>
							<input type="radio" <%=chkIIF(lastDays < 0, modTxt, "")%> name="useyn" value=0 <%=chkIIF(useyn="0", "checked", "")%>> ������
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrap">
		<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
		<input type="button" value="����" onclick="submitContent();" class="cRd1" style="width:100px; height:30px;" />
	</div>
</div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->