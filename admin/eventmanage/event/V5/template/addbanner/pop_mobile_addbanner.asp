<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : pop_mobile_addbanner.asp
' Discription : ����� slide insert
' History : 2016-02-16 ����ȭ
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim eCode , title , eFolder , topimg , btmimg , topaddimg 'floating img
Dim slideimg, menuidx
Dim mode , idx , strSql , sqlStr , sDt , eDt

	eCode = requestCheckvar(request("eC"),16)
	menuidx = requestCheckvar(request("menuidx"),16)
	if menuidx="" or isnull(menuidx) then menuidx=0
	title = "�����̵� ��� �˾�(M)"
	eFolder = eCode

	If eCode <> "" Then
		strSql = "SELECT evt_startdate , evt_enddate " + vbcrlf
		strSql = strSql & " from db_event.dbo.tbl_event where evt_code = '"& eCode &"' " 
		rsget.Open strSql,dbget,1
		IF Not rsget.Eof Then
			sDt		= date()
			eDt		= rsget("evt_enddate")
		End If
		rsget.close()
	End If

%>
<!-- #include virtual="/admin/lib/popheaderslide.asp"-->
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type='text/javascript'>
function chkAllItem() {
	if($("input[name='chkIdx']:first").attr("checked")=="checked") {
		$("input[name='chkIdx']").attr("checked",false);
	} else {
		$("input[name='chkIdx']").attr("checked","checked");
	}
}

function saveList() {
	var chk=0;
	var frm = document.frmList;
	$("form[name='frmList']").find("input[name='chkIdx']").each(function(){
		if($(this).attr("checked")) chk++;
	});
	if(chk==0) {
		alert("�����Ͻ� ���縦 �������ּ���.");
		return;
	}

	if(confirm("�����Ͻ� ����� ���� ������ �����Ͻðڽ��ϱ�?")) {
		frm.action="pop_mobile_addbanner_proc.asp";
		frm.submit();
	}
}

//'������ ����
function slideimgDel(v){
	if (confirm("��ʰ� �����˴ϴ�. ���� �Ͻðڽ��ϱ�?")){
		document.frmdel.chkIdx.value = v;
		document.frmdel.submit();
	}
}

function jsSetImg(sFolder, sImg, sName, sSpan){ 
	var winImg;
	winImg = window.open('/admin/eventmanage/common/pop_event_uploadimgV2.asp?yr=<%=Year(now())%>&sF='+sFolder+'&sImg='+sImg+'&sName='+sName+'&sSpan='+sSpan,'popImg','width=370,height=150');
	winImg.focus();
}

function jsDelImg(sName, sSpan){
	if(confirm("�̹����� �����Ͻðڽ��ϱ�?\n\n���� �� �����ư�� ������ ó���Ϸ�˴ϴ�.")){
	   eval("document.all."+sName).value = "";
	   eval("document.all."+sSpan).style.display = "none";
	}
}
</script>
<script type="text/javascript">
//��ũ������
function showDrop(){
	$(".selectLink ul").show();
}

//�����Է�
function populateTextBox(v){
	var val = v;
	$("#mblink").val(val);
	$("#blink").val(val);
	$(".selectLink ul").css("display","none");
}

function linkcopy(){
	var val = $("#mblink").val();
	$("#blink").attr("value",val);
	$(".selectLink ul").css("display","none");
}

//�޷�
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'popCal','width=250, height=200');
	winCal.focus();
}

function jsPopCal_2(sName,sChkname){
	if (eval("document.all."+sChkname).checked){
		alert("üũ �ڽ� ������ ������ �����մϴ�");
		return false;
	}else{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'popCal','width=250, height=200');
		winCal.focus();
	}
}

function simgsubmit(){
	// ��� ��� 1 row���
	var frm = document.slideimgfrm;
	
	if (!frm.gubun[0].checked&&!frm.gubun[1].checked&&!frm.gubun[2].checked){
		alert("��ġ�� �������ּ���");
		frm.gubun[0].focus();
		return false;
	}

	if (!frm.btitle.value){ alert("Alt���� �Է� ���ּ���");frm.btitle.focus();return false; }
	//if (!frm.bst_date.value){ alert("�������� �Է� ���ּ���");frm.bst_date.focus();return false; }
	//if (!frm.bed_date.value){ alert("�������� �Է� ���ּ���");frm.bed_date.focus();return false; }

	//if(frm.bst_date.value > frm.bed_date.value){ alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���"); frm.bed_date.focus(); return false; }

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

//������ ������ ����
function jscalcopy(s,e,y){
	if (eval("document.all."+y).checked){
		if(confirm("�̺�Ʈ �Ⱓ�� ���� ������ �Ͻðڽ��ϱ�?")){
			eval("document.all."+s).value = "<%=sDt%>";
			eval("document.all."+e).value = "<%=eDt%>";
		}
	}
//	else{
//		if(confirm("��¥�� �ʱ�ȭ �Ͻðڽ��ϱ�?")){
//			eval("document.all."+s).value = "";
//			eval("document.all."+e).value = "";
//		}
//	}
}
function fnMasterInfoSet(){
	var winMasterInfo;
	winMasterInfo = window.open('/admin/eventmanage/event/v5/popup/pop_multicontents_masterinfo.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>','menuinfo','width=1024,height=850,scrollbars=yes,resizable=yes');
	winMasterInfo.focus();
}

$(function(){
    $("#accordion").accordion();
	//�巡��
	$("#subList").sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).find("input[name^='viewidx']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).find("input[name^='viewidx']").each(function(){
				$(this).val(i);
				i++;
			});
		}
	});
});
</script>
</head>
<body>
<div class="slideRegister adminMob bnrRegister">
	<h1>��� ���(MOBILE)&nbsp;&nbsp;/&nbsp;&nbsp;<a href="/admin/eventmanage/event/v5/template/addbanner/pop_pc_addbanner.asp?eC=<%=eCode%>&menuidx=<%=menuidx%>" style="color:#FFFFFF">(PC Web)</a> <input type="button" class="btn" value="���/���̾ƿ� ����" onclick="fnMasterInfoSet();"></h1>
	<div class="register">
		<dl>
			<dd>
				<form name="slideimgfrm" method="post" action="pop_mobile_addbanner_proc.asp" style="margin:0px;">
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<input type="hidden" name="mode" value="SI"/>
				<input type="hidden" name="bimg" value=""/>
				<input type="hidden" name="blink" id="blink" value=""/>
				<input type="hidden" name="sDt" value="<%=sDt%>"/>
				<input type="hidden" name="eDt" value="<%=eDt%>"/>
				<div class="insertImg">
					<table class="tbType1 listTb">
						<colgroup>
							<col width="13%" /><col width="20%" /><col /><col width="28%" /><col width="9%" />
						</colgroup>
						<tbody>
						<tr>
							<td>
								<span><input type="radio" name="gubun" value="1" id="gt"/> <label for="gt">��</label></span>
								<span class="lMar05"><input type="radio" name="gubun" value="2" id="gm"/> <label for="gm">��</label></span>
								<span class="lMar05"><input type="radio" name="gubun" value="3" id="gb"/> <label for="gb">��</label></span>
							</td>
							<td>
								<input class="button" type="button" value="�̹��� �ҷ�����" name="mbimg" onClick="jsSetImg('<%=eFolder%>','','bimg','spanslideimg');"/>
								<div id="spanslideimg"></div>
								<div class="tMar10"><input type="text" name="btitle" placeholder="Alt�� �Է�" /></div>
							</td>
							<td>
								<div class="selectLink">
									<input type="text" value="��ũ�� �Է�(����)" onclick="showDrop();" id="mblink" onkeyup="linkcopy();" />
									<ul style="display:none;">
										<li onclick="populateTextBox('');">���þ���</li>
										<li onclick="populateTextBox('/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�');">/event/eventmain.asp?eventid=�̺�Ʈ�ڵ�</li>
										<li onclick="populateTextBox('#group�׷��ڵ�');">#group�׷��ڵ�</li>
										<li onclick="populateTextBox('#replyList');">#replyList(�ڸ�Ʈ�ٷΰ���)</li>
										<li onclick="populateTextBox('#replyPrdList');">#replyPrdList(��ǰ�ı�ٷΰ���)</li>
										<li onclick="populateTextBox('/category/category_itemprd.asp?itemid=��ǰ�ڵ�');">/category/category_itemprd.asp?itemid=��ǰ�ڵ� (O)</li>
										<li onclick="populateTextBox('/category/category_list.asp?disp=ī�װ���');">/category/category_list.asp?disp=ī�װ���</li>
										<li onclick="populateTextBox('/street/street_brand.asp?makerid=�귣����̵�');">/street/street_brand.asp?makerid=�귣����̵�</li>
									</ul>
								</div>
							</td>
							<!--<td>
								<p>������ : <input type="text" onclick="jsPopCal('bst_date');" value="<%=sDt%>" style="width:85px; cursor:pointer;" name="bst_date" readonly> ~ ������ : <input type="text" value="<%=eDt%>" onclick="jsPopCal('bed_date');" style="width:85px; cursor:hand;" name="bed_date" readonly></p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag" value="Y" onclick="jscalcopy('bst_date','bed_date','bdate_flag');" checked /> �̺�Ʈ �Ⱓ ���� ����</p>
							</td>-->
							<td><input type="button" class="btn" value="���" onclick="simgsubmit();"></td>
						</tr>
						</tbody>
					</table>
				</div>
				</form>

				<form name="frmList" method="POST" action="" style="margin:0;">
				<input type="hidden" name="mode" value="SU"/>
				<input type="hidden" name="device" value="M"/>
				<input type="hidden" name="eventid" value="<%=eCode%>"/>
				<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
				<input type="hidden" name="sDt" value="<%=sDt%>"/>
				<input type="hidden" name="eDt" value="<%=eDt%>"/>
				<div class="tMar20">
					<table class="tbType1 listTb">
						<colgroup>
							<col width="5%" /><col width="13%" /><col width="20%" /><col /><col width="28%" /><col width="9%" />
						</colgroup>
						<thead>
						<tr>
							<th><input type="checkbox" id="" onclick="chkAllItem();"/></th>
							<th>��ġ</th>
							<th>�̹���</th>
							<th>��ũ(����)</th>
							<!--<th>������/������</th>-->
							<th>��뿩��</th>
						</tr>
						</thead>
						<tbody id="subList">
						<% 
							If eCode <> "" Then 

							sqlStr = "SELECT idx, gubun, bimg, btitle, blink, bdate_flag, bst_date, bed_date, isusing, viewidx" & vbcrlf
							sqlStr = sqlStr & " from db_event.dbo.tbl_event_mobile_addbanner where evt_code='"& eCode &"'" & vbcrlf
							sqlStr = sqlStr & " and menuidx=" & menuidx & vbcrlf
							sqlStr = sqlStr & " order by gubun asc, viewidx asc"
							rsget.Open sqlStr,dbget,1
							if Not(rsget.EOF or rsget.BOF) Then
								Do Until rsget.eof
						%>
						<tr class="<%=chkIIF(rsget("isusing")="N","bgGry1","")%>">
							<td><input type="checkbox" name="chkIdx" value="<%=rsget("idx")%>" /></td>
							<td>
								<span><input type="radio" name="gubun<%=rsget("idx")%>" value="1" <%=chkiif(rsget("gubun")=1,"checked","")%> id="gt<%=rsget("idx")%>"/> <label for="gt<%=rsget("idx")%>">��</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="2" <%=chkiif(rsget("gubun")=2,"checked","")%> id="gm<%=rsget("idx")%>"/> <label for="gm<%=rsget("idx")%>">��</label></span>
								<span class="lMar05"><input type="radio" name="gubun<%=rsget("idx")%>" value="3" <%=chkiif(rsget("gubun")=3,"checked","")%> id="gb<%=rsget("idx")%>"/> <label for="gb<%=rsget("idx")%>">��</label></span>
								<br><br><input type="text" name="viewidx" value="<%=rsget("viewidx")%>" style="width:50px;">
							</td>
							<td>
								<input class="button" type="button" value="�̹��� �ҷ�����" name="mbimg<%=rsget("idx")%>" onClick="jsSetImg('<%=eFolder%>','','bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"/>
								<input type="hidden" name="bimg<%=rsget("idx")%>" value="<%=rsget("bimg")%>"/><%' �̹��� %>
								<div id="spanslideimg<%=rsget("idx")%>">
									<img src="<%=rsget("bimg")%>" style="width:100px;" alt="<%=rsget("btitle")%>"/>
									<%IF rsget("bimg") <> "" THEN %>
									<a href="javascript:jsDelImg('bimg<%=rsget("idx")%>','spanslideimg<%=rsget("idx")%>');"><img src="/images/icon_delete2.gif" border="0" class="delImg"></a>
									<%END IF%>
								</div>
								<div class="tMar10"><input type="text" name="btitle<%=rsget("idx")%>" value="<%=rsget("btitle")%>" placeholder="Alt�� �Է�"/></div>
							</td>
							<td><input type="text" name="blink<%=rsget("idx")%>" value="<%=rsget("blink")%>" /></td>
							<!--<td>
								<p>������ : <input type="text" onclick="jsPopCal_2('bst_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bst_date")%>" style="width:85px; cursor:pointer;" name="bst_date<%=rsget("idx")%>"> ~ ������ : <input type="text" onclick="jsPopCal_2('bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');" value="<%=rsget("bed_date")%>" style="width:85px; cursor:pointer;" name="bed_date<%=rsget("idx")%>"></p>
								<p class="tMar05"><input type="checkbox" name="bdate_flag<%=rsget("idx")%>" value="Y" <%=chkiif(rsget("bdate_flag")="Y","checked","")%> onclick="jscalcopy('bst_date<%=rsget("idx")%>','bed_date<%=rsget("idx")%>','bdate_flag<%=rsget("idx")%>');"/> �̺�Ʈ �Ⱓ ���� ����</p>
							</td>-->
							<td>
								<span><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="Y","checked","")%> value="Y"/> Y</span>
								<span class="lMar10"><input type="radio" name="isusing<%=rsget("idx")%>" <%=chkiif(rsget("isusing")="N","checked","")%> value="N"/> N</span>
								<br />
								<input type="button" class="btn tMar05" value="����" onclick="slideimgDel(<%=rsget("idx")%>);"/>
							</td>
						</tr>
						<% 
								rsget.movenext
								Loop
							End If
							rsget.close

							End If
						%>
						</tbody>
					</table>
					<p class="tMar20 ct">
						<input type="button" class="btn" value="��ü ����" onclick="chkAllItem();">
						<input type="button" class="btn" value="���� ����" onClick="saveList();" title="ǥ�ü��� �� ��뿩�θ� �ϰ������մϴ�.">
					</p>
				</div>
				</form>
			</dd>
		</dl>
	</div>
</div>
<form name="frmdel" method="POST" action="pop_mobile_addbanner_proc.asp" style="margin:0px;">
<input type="hidden" name="sDt" value="<%=sDt%>"/>
<input type="hidden" name="eDt" value="<%=eDt%>"/>
<input type="hidden" name="eventid" value="<%=eCode%>"/>
<input type="hidden" name="menuidx" value="<%=menuidx%>"/>
<input type="hidden" name="mode" value="SD"/>
<input type="hidden" name="chkIdx" />
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->