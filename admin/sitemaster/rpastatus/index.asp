<% Option Explicit %>
<%
'###########################################################
' Description : rpa ���� ���� ����Ʈ
' Hieditor : 2021.07.20 ������ ����
'###########################################################

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/rpastatus/rpastatuscls.asp"-->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
Dim loginUserId, i, currpage, pagesize, research, startdate, enddate
Dim rpatype, rpatitle, rpacontents, rparegdate, rpaissuccess
Dim oRpaStatusList

loginUserId = session("ssBctId") '// �α����� ����� ���̵�
currpage = requestcheckvar(request("page"), 20) '// ���� ������ ��ȣ
rpatype = requestcheckvar(request("rpatype"), 240) '// Ÿ�Ը�(cls�� Ÿ�� ���� ����)
research = requestcheckvar(request("research"), 20) '// ��˻�����
startdate = requestcheckvar(request("startdate"), 20) '// ����� ���� �˻���
enddate = requestcheckvar(request("enddate"), 20) '// ����� ���� �˻���
rpaissuccess = requestcheckvar(request("rpaissuccess"), 20) '// ���� ���� ����

If Trim(currpage)="" Then
	currpage = "1"
End If
pagesize = 30


'// ����Ʈ�� �����´�.
set oRpaStatusList = new CgetRpaStatus
	oRpaStatusList.FRectcurrpage = currpage
	oRpaStatusList.FRectpagesize = pagesize
	If Trim(research)="on" Then
		oRpaStatusList.FRectType        = rpatype
		oRpaStatusList.FRectIsSuccess   = rpaissuccess
		oRpaStatusList.FRectStartdate   = startdate
		oRpaStatusList.FRectEnddate     = enddate
	End If
    oRpaStatusList.GetHalfDeliveryPayList()
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script language="JavaScript" src="/js/xl.js"></script>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="JavaScript" src="/js/calendar.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>
document.domain = "10x10.co.kr";

function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopMenuEdit(menupos){
	var popwin = window.open('/admin/menu/pop_menu_edit.asp?mid=' + menupos,'PopMenuEdit','width=600, height=400, scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnMenuFavoriteAct(mode) {
	var frm = document.frmMenuFavorite;
	frm.mode.value = mode;
	var msg;
	var ret;
	if (mode == "delonefavorite") {
		msg = "���ã�⿡�� �����Ͻðڽ��ϱ�?";
	} else {
		msg = "���ã�⿡ �߰��Ͻðڽ��ϱ�?";
	}
	ret = confirm(msg);
	if (ret) {
		frm.submit();
	}
}

function goPage(page){
	<% if trim(research)="on" then %>
	    location.href='?page=' + page + '&research=on&menupos=<%=request("menupos")%>&rpatype=<%=rpatype%>&startdate=<%=startdate%>&enddate=<%=enddate%>&rpaissuccess=<%=rpaissuccess%>';
	<% else %>
	    location.href="?page=" + page;
	<% end if %>
}

function goSearchRpaStatus()
{
	document.frm1.submit();
}

function jsChkAll(){
var frm;
frm = document.frm;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkidx) !="undefined"){
	   	   if(!frm.chkidx.length){
		   	 	frm.chkidx.checked = true;
		   }else{
				for(i=0;i<frm.chkidx.length;i++){
					frm.chkidx[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkidx) !="undefined"){
	  	if(!frm.chkidx.length){
	   	 	frm.chkidx.checked = false;
	   	}else{
			for(i=0;i<frm.chkidx.length;i++){
				frm.chkidx[i].checked = false;
			}
		}
	  }

	}
}

function goIsUsingModifyAll(tp) {
	var itemcount = 0;
	var frm;
	var ck=0;
	frm = document.frm;
	if(typeof(frm.chkidx) !="undefined"){
		if(!frm.chkidx.length){
			if(!frm.chkidx.checked){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
			//frm.itemidarr.value = frm.chkitem.value;
			//frm.itemdataarr.value = frm.viewitemdata.value;
		}else{
			//frm.itemidarr.value = "";
			for(i=0;i<frm.chkidx.length;i++){
				if(frm.chkidx[i].checked) {
					ck=ck+1;
					if (frm.itemisusingarr.value==""){
						frm.itemisusingarr.value =  frm.chkidx[i].value;
					}else{
						frm.itemisusingarr.value = frm.itemisusingarr.value + "," +frm.chkidx[i].value;
					}
				}
			}

			if (frm.itemisusingarr.value == ""){
				alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
				return;
			}
		}
	}else{
		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
		return;
	}

	$("#isusingtype").val(tp);
	if(confirm("�����Ͻ� ��� ��ǰ�� ��뿩�ΰ� ����˴ϴ�.\n�����Ͻðڽ��ϱ�?")) {
		document.frm.submit();
	} else {
		return false;
	}
}

function jsEtcSaleMarginJungsan(makerid){
	var upfrm1 = document.frmEtcJOne;
    upfrm1.makerid.value=makerid;

    if (confirm("�ۼ� �Ͻðڽ��ϱ�?")){
        upfrm1.submit();
    }
}

</script>
<div class="">
	<%' ��� �˻��� ���� %>
	<form name="frm1" id="frm1" method="get" action="/admin/sitemaster/rpastatus/index.asp">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<%' search %>
	<div class="searchWrap" style="border-top:none;">
		<div class="search">
			<ul>
				<li>
					<label class="formTit">������ :</label>
                    <select name="rpatype">
                        <option value="">����</option>			
                        <option value="���̹�����" <%IF trim(rpatype)="���̹�����" THEN%>selected<%END IF%>>���̹����� ���곻�� �ٿ�ε�</option>
                        <option value="�̼���" <%IF trim(rpatype)="�̼���" THEN%>selected<%END IF%>>�̼��� ���ڰ�꼭 �ٿ�ε�</option>
                        <option value="KICC����" <%IF trim(rpatype)="KICC����" THEN%>selected<%END IF%>>KICC ���γ��� �ٿ�ε�</option>
                        <option value="KICC�Ա�" <%IF trim(rpatype)="KICC�Ա�" THEN%>selected<%END IF%>>KICC �Աݳ��� �ٿ�ε�</option>
                        <option value="���޸�����" <%IF trim(rpatype)="���޸�����" THEN%>selected<%END IF%>>���޸� ���곻�� �ٿ�ε�(����)</option>
                        <option value="���޻����" <%IF trim(rpatype)="���޻����" THEN%>selected<%END IF%>>���޻� ���� ���� �� ����</option>
                        <option value="�������" <%IF trim(rpatype)="�������" THEN%>selected<%END IF%>>�������</option>
                        <option value="īī������Ʈ�ɼ�" <%IF trim(rpatype)="īī������Ʈ�ɼ�" THEN%>selected<%END IF%>>īī�� ����Ʈ �ɼ� ��� ��Ī</option>
                        <option value="����ī��" <%IF trim(rpatype)="����ī��" THEN%>selected<%END IF%>>����ī�� SCM ���ε�</option>
                        <option value="�����" <%IF trim(rpatype)="�����" THEN%>selected<%END IF%>>����� ���ǻ��� ����</option>
                        <option value="���޸��ֹ�" <%IF trim(rpatype)="���޸��ֹ�" THEN%>selected<%END IF%>>���޸� �ֹ� ����</option>
                        <option value="���������" <%IF trim(rpatype)="���������" THEN%>selected<%END IF%>>������� ����۾�</option>
                    </select>		
				</li>
				<li>
					<p class="formTit">�Ⱓ</p>
					<input type="text" id="startdate" name="startdate" value="<%=startdate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "startdate", trigger    : "startdate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>
                     ~ 
					<p class="formTit"></p>
					<input type="text" id="enddate" name="enddate" value="<%=enddate%>" class="formTxt" size="10" maxlength="10" style="margin-bottom:13px;"/>
					<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" style="vertical-align:middle;"/>
					<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "enddate", trigger    : "enddate_trigger",
							onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					</script>                     
				</li>
				<li>
					<p class="formTit">�������� :</p>
					<select class="formSlt" id="rpaissuccess" name="rpaissuccess" title="�������� ����">
						<option value="" <% If rpaissuccess = "" Then %> selected <% End If %>>��ü</option>
						<option value="1" <% If rpaissuccess = "1" Then %> selected <% End If %>>����</option>
						<option value="0" <% If rpaissuccess = "0" Then %> selected <% End If %>>����</option>
					</select>
				</li>
			</ul>
		</div>
		<input type="button" class="schBtn" value="�˻�" onclick="goSearchRpaStatus();" />
	</div>
	<!-- //search -->
	</form>

	<div class="cont">
		<div class="pad20">
			<div class="pieceList">
				<div class="rt bPad10 rPad10">
					<p class="totalNum">�� ��ϼ� : <strong><%=FormatNumber(oRpaStatusList.FtotalCount, 0)%></strong></p>
				</div>
				<div class="tbListWrap">
					<ul class="thDataList">
						<li>
							<p style="width:80px">��ȣ(idx)</p>
							<p style="width:100px">������</p>
                            <p style="width:450px">����</p>
							<!--p style="width:600px">����</p-->
                            <p style="width:80px">��������</p>
							<p style="width:90px">�����</p>
							<p style="width:150px"></p>
						</li>
					</ul>
					<ul id="sortable" class="tbDataList">
						<% If oRpaStatusList.FResultcount > 0 Then %>
							<% For i=0 To oRpaStatusList.Fresultcount-1 %>
							<% If oRpaStatusList.FrpaStatusList(i).FisSuccess = 0 Then %>
								<li style="background-color:#FFEDED">
							<% Else %>
								<li style="background-color:#F7FFE6">
							<% End If %>
								<p style="width:80px"><%=oRpaStatusList.FrpaStatusList(i).Fidx%></p>
								<p style="width:100px"><%=getRpaTypeName(oRpaStatusList.FrpaStatusList(i).Ftype)%></p>
								<p style="width:450px" align="left"><%=oRpaStatusList.FrpaStatusList(i).Ftitle%></p>
								<!--p style="text-align:left;width:600px;white-space:pre-line;"><%'replace(oRpaStatusList.FrpaStatusList(i).Fcontents,chr(13)&chr(10),"<br>")%></p-->
								<p style="width:80px"><%=getRpaIsSuccessName(oRpaStatusList.FrpaStatusList(i).FisSuccess)%></p>
								<p style="width:90px"><%=oRpaStatusList.FrpaStatusList(i).Fregdate%></p>
								<p style="width:150px"><button onclick="window.open('popviewrpastatus.asp?idx=<%=oRpaStatusList.FrpaStatusList(i).Fidx%>',null,'height=800,width=1000,status=yes,toolbar=no,menubar=no,location=no');return false;">����Ȯ��</button></p>
							</li>
							<% Next %>
						<% End If %>
					</ul>
					<div class="ct tPad20 cBk1">
						<%=fnDisplayPaging_New2017(currpage, oRpaStatusList.FtotalCount, pagesize, 10, "goPage") %>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script>
$(function() {
	$(".btnOdrChg").on('click',function() {
		if ($("#sortable").hasClass('sortable')) {
			$("#sortable").removeClass('sortable');
			$("#sortable li p:first-child").html("901"); //����Ʈ index�� ���Բ�
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable("destroy");
			$(".btnOdrChg").attr("value", "��������");
			//$(".btnOdrChg").prop("disabled", true); //�˻����� ����� �������� ��ư ��Ȱ��ȭ
			$(".btnRegist").prop("disabled", false);
			$(".infoTxt").hide();
		} else {
			$("#sortable").addClass('sortable');
			$("#sortable li p:first-child").html("<img src='/images/ico_odrchg.png' alt='��������' />");
			$("#sortable li.ui-state-disabled p:first-child").html("����");
			$("#sortable").sortable({
				placeholder:"handling",
				items:"li:not(.ui-state-disabled)"
			}).disableSelection();
			$(".btnOdrChg").attr("value", "����Ϸ�");
			//$(".btnOdrChg").prop("disabled", false);
			$(".btnRegist").prop("disabled", true);
			$(".infoTxt").show();
		}
	});

	$(".memEdit").on('click',function() {
		$(".dimmed").show();
		$(".lyrBox").show();
	});
});
</script>

</body>
</html>
<%
	Set oRpaStatusList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
