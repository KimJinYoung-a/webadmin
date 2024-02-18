<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : popMainContentsEdit.asp
' Discription : ����� ����Ʈ ���� ������ �ۼ�/����
' History : 2010.02.23 ������ ����
'           2012.02.14 ������ - �̴ϴ޷� ��ü
'           2012.12.14 ����ȭ - alt �� �߰�
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/mobile/mo_catetoryMainManageCls.asp" -->
<%
Dim makerid, eC, sqlStr, imgURL
dim idx, poscode, reload
	idx = request("idx")
	poscode = request("poscode")
	reload = request("reload")
	eC = request("eC")
	if idx="" then idx=0

	if reload="on" then
	    response.write "<script>opener.location.reload(); window.close();</script>"
	    dbget.close()	:	response.End
	end if

	dim oMainContents
		set oMainContents = new CMainContents
		oMainContents.FRectIdx = idx
		oMainContents.GetOneMainContents
		imgURL = oMainContents.FOneItem.Fmain_image

dim orderidx
	if oMainContents.FOneItem.fview_order = "" then
		orderidx = 99
	else
		orderidx = oMainContents.FOneItem.fview_order
	end if

dim dateOption
dateOption = request("dateoption")	

if dateOption = "" then
	dateOption = date
end if

%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<title></title>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="http://m.10x10.co.kr/lib/css/main.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<style type="text/css">
.text-bnr .headline {padding:0; background-color:transparent; border:none; color:#0d0d0d;}
.text-bnr .thumbnail img {width:100%;}
</style>
<script src="http://code.jquery.com/jquery-latest.min.js"></script>
<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
	$(function(){
		initiateDate();
	})

	function initiateDate(){
		var date = '<%=dateOption%>'
		var startDateEle = document.getElementById('start_date');
		var endDateEle = document.getElementById('end_date');

		if(date != '' && startDateEle.value == '' && endDateEle.value == '' ){		
			startDateEle.value = date;
			endDateEle.value = date;
		}		
	}	

	function SaveMainContents(frm){

	    if (frm.catecode.value.length<1){
	        alert('ī�װ��� ���� ���� �ϼ���.');
	        frm.catecode.focus();
	        return;
	    }

        if (frm.imgURL.value.length<1){
	        alert('�̹����� ���ε� �� �ּ���.');
	        frm.imgURL.focus();
	        return;
	    }

        if (frm.main_text.value.length<1){
	        alert('���� ī�Ǹ� �Է� �ϼ���.');
	        frm.main_text.focus();
	        return;
	    }

	    if (frm.banner_link.value.length<1){
	        alert('��ũ ���� �Է� �ϼ���.');
	        frm.banner_link.focus();
	        return;
	    }

	    if (frm.start_date.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    if (frm.end_date.value.length!=10){
	        alert('�������� �Է�  �ϼ���.');
	        return;
	    }

	    var vstartdate = new Date(frm.start_date.value.substr(0,4), (1*frm.start_date.value.substr(5,2))-1, frm.start_date.value.substr(8,2));
	    var venddate = new Date(frm.end_date.value.substr(0,4), (1*frm.end_date.value.substr(5,2))-1, frm.end_date.value.substr(8,2));

	    if (vstartdate>venddate){
	        alert('�������� �����Ϻ��� ������ �ȵ˴ϴ�.');
	        return;
	    }

	    if (confirm('���� �Ͻðڽ��ϱ�?')){
	        frm.submit();
	    }
	}

	function ChangeLinktype(comp){
	    if (comp.value=="M"){
	       document.all.link_M.style.display = "";
	       document.all.link_L.style.display = "none";
	    }else{
	       document.all.link_M.style.display = "none";
	       document.all.link_L.style.display = "";
	    }
	}

	function ChangeGubun(comp){
	    location.href = "?poscode=" + comp.value;
	    // nothing;
	}

	// poscode 2071 �߰� �۾�
	function chkopt(v){
		if (v == "2"){
			document.getElementById("culopt").style.display = "";
			document.getElementById("playopt").style.display = "none"; //������� �ּ�
			document.getElementById("callcontents").style.display = "";
		}else if (v == "3"){
			document.getElementById("culopt").style.display = "none";
			document.getElementById("playopt").style.display = ""; //������� �ּ�
			document.frmcontents.maincopy.value = "PLAYing";
		}else{
			document.getElementById("culopt").style.display = "none";
			document.getElementById("playopt").style.display = "none"; //������� �ּ�
			document.frmcontents.maincopy.value = "HITCHHIKER";
		}
	}
</script>
<script type="text/javascript">
	function fileInfo(f){
		var file = f.files; // files �� ����ϸ� ������ ������ �� �� ����

		var reader = new FileReader(); // FileReader ��ü ���
		reader.onload = function(rst){ // �̹����� ������ �ε��� �Ϸ�Ǹ� ����� �κ�
			$('#img_box').empty().html('<img src="' + rst.target.result + '">'); // append �޼ҵ带 ����ؼ� �̹��� �߰�
			// �̹����� base64 ���ڿ��� �߰�
			// �� ����� �����ϸ� ������ �̹����� �̸����� �� �� ����
            $("#imgURL").val(rst.target.result);
		}
		reader.readAsDataURL(file[0]); // ������ �д´�, �迭�̱� ������ 0 ���� ����
	}

	// typing
	function textclone(k,v){
		var frmtext = $("#"+k);
		frmtext.bind("keyup",function(){
			var msg = $(this).val();
			if ($(this).val().length > 0){
				msg = msg.replace(/(?:\r\n|\r|\n)/g, '<br>');
				$("#"+v).html(msg);
			}else{
				$("#"+v).html("");
			}
		});
	}

	$(function() {
		$('input:radio[name="salediv"]').click(function(){
			if($('input:radio[name="salediv"]:checked').val()==1)
			{
				$("#saleinfo1").show();
				//$("#saleinfo2").hide();
			}
			else
			{
				$("#saleinfo1").hide();
				alert('�̺�Ʈ �ڵ带 �Է� ���ּ���');
				$("#saleinfo2").focus();
				//$("#saleinfo2").show();
			}
		});
		<% if eC <> "" then %>
		chkopt("2");
		<% end if %>
	});

	function cultureloadpop(){
		winLast = window.open('/admin/sitemaster/lib/pop_culturelist.asp?gubun=MC&poscode=<%=poscode%>&pidx=<%=idx%>','pLast','width=1200,height=600, scrollbars=yes')
		winLast.focus();
	}
</script>
</head>
<body>
<div class="popWinV17">
	<h1>���</h1>
	<form name="frmcontents" method="post" action="<%=uploadUrl%>/linkweb/doMobileCateRollingBannerReg.asp" onsubmit="return false;" enctype="multipart/form-data">
		<div class="popContainerV17 pad30">
			<div class="ftLt col6">
				<table class="tbType1 writeTb">
					<tr>
						<th width="160">Idx</th>
						<td>
							<% if oMainContents.FOneItem.Fidx<>"" then %>
								<%= oMainContents.FOneItem.Fidx %>
								<input type="hidden" name="idx" value="<%= oMainContents.FOneItem.Fidx %>">
							<% else %>
							<% end if %>
						</td>
					</tr>
					<tr>
						<th width="160">ī�װ�</th>
						<td>
							<% DrawSelectBoxDispCateLarge "catecode", oMainContents.FOneItem.Fcatecode, "" %>
						</td>
					</tr>
					<tr>
						<th width="160">�켱����</th>
						<td>
							<input type="text" name="orderidx" class="formTxt" size=5 value="<%= oMainContents.FOneItem.Fview_order %>">
						</td>
					</tr>
					<tr>
						<th width="160">�̹���</th>
						<td>
							<input type="file" name="file1" value="" size="32" maxlength="32" class="formFile" accept="image/*" onchange="fileInfo(this)">
							<input type="hidden" name="imgURL" id="imgURL" value="<%= imgURL %>">
							<% if oMainContents.FOneItem.Fidx<>"" then %>
							<br>
								<% If imgURL = "" Then %>
									<img src="/images/admin_login_logo2.png" alt="" />
								<% else%>
									<img src="<%= imgURL %>" width="500" alt="" />
								<br> <%= imgURL %>
								<% End If %>
							<% end if %>
						</td>
					</tr>
					<tr>
					  <th width="160">����ī��</th>
					  <td><input type="text" name="main_text" value="<%=oMainContents.FOneItem.Fmain_text%>" class="formTxt" style="width:94%" maxlength="128"> </td>
					</tr>
					<tr>
						<th width="160">��ũ��</th>
						<td>
							<input type="text" class="formTxt" name="banner_link" value="<%= oMainContents.FOneItem.Fbanner_link %>" maxlength="128" style="width:94%">
						</td>
					</tr>
					<tr>
						<th width="160">�ݿ�������</th>
						<td>
							<span class="rMar10">
							<input id="start_date" name="start_date" value="<%=Left(oMainContents.FOneItem.Fstart_date,10)%>" class="formTxt" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="start_date_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
								<input type="text" class="formTxt" name="dummy0" value="00:00:00" size="8" readonly />
							</span>
						</td>
					</tr>
					<tr>
						<th width="160">�ݿ�������</th>
						<td>
							<span class="rMar10"><input id="end_date" name="end_date" value="<%=Left(oMainContents.FOneItem.Fend_date,10)%>" class="formTxt"  size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="end_date_trigger" border="0" style="cursor:pointer; vertical-align:top; margin-left:5px;" /></span>
							<span class="rMar10">
								<input type="text" class="formTxt" name="dummy1" value="23:59:59" size="8" readonly />
							</span>
							<script type="text/javascript">
								var CAL_Start = new Calendar({
									inputField : "start_date", trigger    : "start_date_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_End.args.min = date;
										CAL_End.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
								var CAL_End = new Calendar({
									inputField : "end_date", trigger    : "end_date_trigger",
									onSelect: function() {
										var date = Calendar.intToDate(this.selection.get());
										CAL_Start.args.max = date;
										CAL_Start.redraw();
										this.hide();
									}, bottomBar: true, dateFormat: "%Y-%m-%d"
								});
							</script>
						</td>
					</tr>
					<tr>
						<th width="160">�����</th>
						<td>
							<%= oMainContents.FOneItem.Fregdate %> (<%= oMainContents.FOneItem.Freguserid %>)
						</td>
					</tr>
					<tr>
						<th width="160">��뿩��</th>
						<td>
							<% if oMainContents.FOneItem.Fview_yn="0" then %>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="1">�����</span>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="0" checked >������</span>
							<% else %>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="1" checked >�����</span>
							<span class="rMar10"><input type="radio" name="view_yn" class="formRadio" value="0">������</span>
							<% end if %>
						</td>
					</tr>
				</table>
			</div>
			<div style="position:fixed;left:62%;top:70px;">
				<div class="lPad30 vTop" id="today_preview">
					<%'Ÿ�Ժ� ���ø� %>
					<%'rolling image%>
					<div class="text-bnr">
					<section class="" style="width:375px;">
						<div class="thumbnail" id="img_box">
							<% If imgURL="" Then %>
							<img src="/images/admin_login_logo2.png" alt="" />
							<% Else %>
							<img src="<%=imgURL%>" alt="" />
							<% End If %>
						</div>
					</section>
					</div>
				</div>
			</div>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="����" onClick="SaveMainContents(frmcontents);" class="cRd1" style="width:100px; height:30px;">
		</div>
	</form>
</div>
<%
set oMainContents = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->