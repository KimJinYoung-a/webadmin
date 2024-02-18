<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

If not (Request.ServerVariables("REMOTE_ADDR") = "61.252.133.75" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.105" or Request.ServerVariables("REMOTE_ADDR") = "61.252.133.106") Then
	Response.End
End If
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/search/search_manageCls.asp"-->
<%
Dim i, cCurator, vIdx, vTitle, vViewGubun, vRegUserName, vSDate, vEDate, vRegdate, vLastUserName, vLastdate, vMemo, vUseYN
Dim vKwArr, vUnitArr, vUnit, vUnitCount, vShhmmss, vEhhmmss
vIdx = requestCheckVar(Request("idx"),15)

If vIdx <> "" Then
	Set cCurator = New CSearchMng
	cCurator.FRectIdx = vIdx
	cCurator.sbCuratorDetail

	vTitle = cCurator.FOneItem.Ftitle
	vViewGubun = cCurator.FOneItem.Fviewgubun
	vSDate = cCurator.FOneItem.Fsdate
	vEDate = cCurator.FOneItem.Fedate
	vShhmmss = cCurator.FOneItem.Fshhmmss
	vEhhmmss = cCurator.FOneItem.Fehhmmss
	vRegUserName = cCurator.FOneItem.Fregusername
	vRegdate = cCurator.FOneItem.Fregdate
	vLastUserName = cCurator.FOneItem.Flastusername
	vLastdate = cCurator.FOneItem.Flastdate
	vMemo = cCurator.FOneItem.Fmemo
	vUseYN = cCurator.FOneItem.Fuseyn
	vUnitArr = cCurator.FUnitArr

	
	If IsArray(cCurator.FKeywordArr) Then
		For i =0 To UBound(cCurator.FKeywordArr,2)
			If i = 0 Then
				vKwArr = cCurator.FKeywordArr(0,i)
			Else
				vKwArr = vKwArr & "," & cCurator.FKeywordArr(0,i)
			End If
		Next
	End If
	Set cCurator = Nothing
Else
	vViewGubun = "period"
	vUseYN = "y"
	vUnitCount = 0
	vShhmmss = "10:00:00"
	vEhhmmss = "09:59:59"
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<script language='javascript'>
function jsCuratorUnit(i){
	var popcuunitreg;
	popcuunitreg = window.open('keywordQratingUnit.asp?idx='+i+'','popcuunitreg','width=1500,height=900,scrollbars=yes,resizable=yes');
	popcuunitreg.focus();
}

function jsCuratorSave(){
	var msg;
	msg = "";
	
	if($("#title").val() == ""){
		alert("������ �Է��ϼ���.");
		return;
	}
	if($("#keyword").val() == ""){
		alert("�˻� Ű���带 �Է��ϼ���.");
		return;
	}
	if($("#sdate").val() == "" || $("#edate").val() == ""){
		alert("������, �������� �Է����ּ���.");
		return;
	}

	<% If vIdx <> "" Then %>
		if($("#unit").val() == ""){
			alert("Unit�� 4~10 �� ���̷� �Է��ϼ���.");
			return;
		}
		if($("#unitcount").val() < 4 || $("#unitcount").val() > 10){
			msg = "������(Unit)�� 4~10���� �ƴ� ��� �ڵ����� ���������� ����˴ϴ�.\n";
		}
		if(confirm("" + msg + "�����Ͻðڽ��ϱ�?") == true) {
			if(msg != ""){
				$("input:radio[name='useyn']:radio[value='n']").attr("checked",true);
			}
			frm1.submit();
	     } else {
	     	return false;
	     }
	<% Else %>
		//msg = "���� �� Unit ������ ����ؾ� ���� ������ �˴ϴ�.\n";
		frm1.submit();
	<% End If %>
}

function jsUnitDelete(g,i){
	$("#unitgubun").val(g);
	$("#unitcontentsidx").val(i);
	frm2.submit();
}
</script>
</head>
<body>
<div id="calendarPopup" style="position: absolute; visibility: hidden; z-index: 2;"></div>
<div class="contSectFix scrl">
	<form name="frm1" action="keywordQratingProc.asp" method="post" style="margin:0px;" target="iframeproc">
	<input type="hidden" name="idx" value="<%=vIdx%>">
	<div class="cont">
		<div class="searchWrap inputWrap">
			<h3>- �⺻ ����</h3>
			<table class="writeTb tMar10">
				<colgroup>
					<col width="14%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>���� *</div></th>
					<td><input type="text" class="formTxt" id="title" name="title" value="<%=vTitle%>" maxlength="10" placeholder="10�� �̳��� Ű���� ť������ ������ �Է����ּ���." style="width:50%" /></td>
				</tr>
				<tr>
					<th><div>�˻� Ű���� *</div></th>
					<td>
						<input type="text" class="formTxt" id="keyword" name="keyword" value="<%=vKwArr%>" placeholder="Ű���� ť�����͸� ������ �˻� Ű���带 ',(��ǥ)' �������� �Է����ּ���." style="width:99%" maxlength="200" />
						<input type="hidden" id="keyword_in_db" name="keyword_in_db" value="<%=vKwArr%>">
					</td>
				</tr>
				<tr>
					<th><div>���� �Ⱓ *</div></th>
					<td>
						<span><input type="hidden" id="termSet" name="viewgubun" value="<%=vViewGubun%>" /></span>
						<span>
							<input type="text" class="formTxt" id="sdate" name="sdate" value="<%=vSDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="sdate_trigger" alt="�޷����� �˻�" />
							<script language="javascript">
								var CAL_Start = new Calendar({
									inputField : "sdate", trigger    : "sdate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="shhmmss" name="shhmmss" value="<%=vShhmmss%>" style="width:60px" maxlength="8" readonly />
							~
							<input type="text" class="formTxt" id="edate" name="edate" value="<%=vEDate%>" style="width:100px" placeholder="������" maxlength="10" readonly />
							<img src="/images/admin_calendar.png" id="edate_trigger" alt="�޷����� �˻�" />
							<script language="javascript">
								var CAL_End = new Calendar({
									inputField : "edate", trigger    : "edate_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_Start.args.max = date;
										CAL_Start.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
							<input type="text" class="formTxt" id="ehhmmss" name="ehhmmss" value="<%=vEhhmmss%>" style="width:60px" maxlength="8" readonly />
						</span>
					</td>
				</tr>
				<tr>
					<th><div>��� ���� *</div></th>
					<td>
						<span class="rMar10"><input type="radio" id="useyny" name="useyn" value="y" <%=CHKIIF(vUseYN="y","checked","")%> /> <label for="useyny">�����</label></span>
						<span class="rMar10"><input type="radio" id="useynn" name="useyn" value="n" <%=CHKIIF(vUseYN="n","checked","")%> /> <label for="useynn">������</label></span>
					</td>
				</tr>
				<tr>
					<th><div>���</div></th>
					<td><textarea class="formTxtA" rows="6" style="width:99%;" id="memo" name="memo"><%=vMemo%></textarea></td>
				</tr>
				</tbody>
			</table>
		</div>
		<% If vIdx <> "" Then %>
		<div class="pad20">
			<h3>- Unit ����</h3>
			<div class="tPad20">
				<input type="button" class="btn" value=" ������ �˻� " onClick="jsCuratorUnit('<%=vIdx%>');" />
				 * Unit�� 4�� �̸��� ��� �ڵ����� Ű���� ť�����͸� <span class="cRd1">������</span> ó���մϴ�. (�̺�Ʈ �������� Ȯ���Ͽ� ����ּ���.)
			</div>
			<div id="unitlist">
				<table class="tbType1 listTb tMar10">
					<thead>
					<tr>
						<th><div>����</div></th>
						<th><div>Unit</div></th>
						<th><div>Unit��</div></th>
						<th><div>������</div></th>
						<th><div>����</div></th>
					</tr>
					</thead>
					<tbody id="unitinsertlist">
					<%
					If IsArray(vUnitArr) Then
						For i =0 To UBound(vUnitArr,2)
							'vUnit : ex) event$67890$1
							If i = 0 Then
								vUnit = vUnitArr(1,i) & "$" & vUnitArr(2,i) & "$" & vUnitArr(3,i)
							Else
								vUnit = vUnit & "," & vUnitArr(1,i) & "$" & vUnitArr(2,i) & "$" & vUnitArr(3,i)
							End If
					%>
							<tr>
								<td><%=(i+1)%></td>
								<td><%=vUnitArr(1,i)%></td>
								<td class="lt"><%=db2html(vUnitArr(0,i))%></td>
								<td>
									<%
										If vUnitArr(1,i) = "event" Then
											If date() <= vUnitArr(4,i) Then
												vUnitCount = vUnitCount + 1
											Else
												Response.Write "<font color=red>[����]</font> "
											End If
											Response.Write Left(vUnitArr(4,i),10)
										Else
											vUnitCount = vUnitCount + 1
										End If
									%>
								</td>
								<td><input type="button" class="btn" value="����" onClick="jsUnitDelete('<%=vUnitArr(1,i)%>','<%=vUnitArr(2,i)%>');" /></td>
							</tr>
					<%
						Next
					End IF
					%>
					</tfoot>
				</table>
				<input type="hidden" id="unit" name="unit" value="<%=vUnit%>">
				<input type="hidden" id="unitcount" name="unitcount" value="<%=vUnitCount%>">
				<div class="tPad20 rt">
					 * Unit ���� <span class="cRd1" id="unitcountspan"><%=vUnitCount%></span> �� (����� �̺�Ʈ�� ī��Ʈ ���� �ʽ��ϴ�.)
				</div>
			</div>
			<input type="hidden" id="unit_in_db" name="unit_in_db" value="<%=vUnit%>">
			<div class="tMar20 ct">
				<input type="button" value="���" onclick="jsCuratorSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="���" onclick="location.href='keywordQratingManageList.asp';" style="width:100px; height:30px;" />
			</div>
		</div>
		<% Else %>
		<div class="pad20">
			<div class=" ct">
				<input type="button" value="����" onclick="jsCuratorSave();" class="cRd1" style="width:100px; height:30px;" />
				<input type="button" value="���" onclick="location.href='keywordQratingManageList.asp';" style="width:100px; height:30px;" />
			</div>
		</div>
		<% End If %>
	</div>
	</form>
</div>
<form name="frm2" action="keywordQratingProc.asp" method="post" target="iframeproc" style="margin:0px;">
<input type="hidden" id="action" name="action" value="unitdelete">
<input type="hidden" id="idx" name="idx" value="<%=vIdx%>">
<input type="hidden" id="unitgubun" name="unitgubun" value="">
<input type="hidden" id="unitcontentsidx" name="unitcontentsidx" value="">
</form>
<iframe src="about:blank" name="iframeproc" width="0" height="0" frameborder="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->