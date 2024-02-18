<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemaster/mailzineCodeCls.asp"-->
<%
Dim arrList,intLoop
Dim mailzineKind, contentsKind, idx
Dim contentsEa, contentsCode, kindCode
Dim clsCode, sMode

contentsKind = requestCheckVar(Request("contentsKind"),10)
mailzineKind = requestCheckVar(Request("mailzineKind"),10)
idx   = requestCheckVar(Request("idx"),10)
sMode ="I"

Set clsCode = new CEventCommonCode  	
IF idx > "0" THEN
	sMode ="U"
	clsCode.FRectIDX  = idx 
	clsCode.fnGetTemplateCont
	kindCode = clsCode.FkindCode
	contentsCode = clsCode.FcontentsCode
	contentsEa  = clsCode.FcontentsEa
END IF		
clsCode.FRectkindCode = mailzineKind
arrList = clsCode.fnGetTemplateList
Set clsCode = nothing

%>
<link rel="stylesheet" type="text/css" href="/admin/eventmanage/event/v5/lib/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/admin/eventmanage/event/v5/lib/css/adminCommon.css" />
<link rel="stylesheet" href="https://cdn.materialdesignicons.com/3.6.95/css/materialdesignicons.min.css">
<style>
html {overflow-y:auto;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script language="javascript">
<!--
	// 코드타입 변경이동
	function jsSetCode(contentsKind,mailzineKind,idx){	
		self.location.href = "popManageTemplate.asp?contentsKind="+contentsKind+"&mailzineKind="+mailzineKind+"&idx="+idx;
	}
	
	//코드 검색
	function jsSearch(mailzineKind){
		self.location.href = "popManageTemplate.asp?mailzineKind="+mailzineKind;
	}

	function jsCodeSort(){
		document.frmSearch.action="procTemplate.asp";
		document.frmSearch.submit();
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.mailzineKind.value) {
			alert("메일진 종류를 선택해 주세요");
			frm.mailzineKind.focus();
			return false;
		}
			 
		if(!frm.contentsKind.value) {
			alert("컨텐츠 종류를 선택해 주세요");
			frm.contentsKind.focus();
			return false;
		}
		return true;
	}

	function jsDeleteCode(idx){
		if(confirm("삭제하시겠습니까?")){
			document.frmReg.mode.value="D";
			document.frmReg.idx.value=idx;
			document.frmReg.submit();
		}
	}

	function fnContentsEaSet(objval){
		if(objval==20 || objval==21 || objval==22 || objval==23){
			$("#contentsEA").html("<option value='1'>1개</option>");
		}
		else if(objval==24){
			$("#contentsEA").html("<option value='1'>1개</option>");
		}
		else if(objval==25){
			$("#contentsEA").html("<option value='1'>1개</option>");
		}
		else if(objval==26){
			$("#contentsEA").html("<option value='4'>4개</option><option value='8'>8개</option>");
		}
		else if(objval==27 || objval==28 || objval==29){
			$("#contentsEA").html("<option value='3'>3개</option><option value='6'>6개</option><option value='9'>9개</option><option value='12'>12개</option><option value='15'>15개</option>");
		}
		else if(objval==30 || objval==31){
			$("#contentsEA").html("<option value='1'>1개</option><option value='3'>3개</option>");
		}
	}

	$(function(){
		$("#accordion").accordion();
		//드래그
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
		<% if contentsCode="20" or contentsCode="21" or contentsCode="22" or contentsCode="23" then %>
			$("#contentsEA").html("<option value='1' selected>1개</option>");
		<% elseif contentsCode="24" then %>
			$("#contentsEA").html("<option value='1' selected>1개</option>");
		<% elseif contentsCode="25" then %>
			$("#contentsEA").html("<option value='1' selected>1개</option>");
		<% elseif contentsCode="26" then %>
			<% if contentsEa=4 then %>
				$("#contentsEA").html("<option value='4' selected>4개</option><option value='8'>8개</option>");
			<% else %>
				$("#contentsEA").html("<option value='4'>4개</option><option value='8' selected>8개</option>");
			<% end if %>
		<% elseif contentsCode="27" or contentsCode="28" or contentsCode="29" then %>
			<% if contentsEa=3 then %>
				$("#contentsEA").html("<option value='3' selected>3개</option><option value='6'>6개</option><option value='9'>9개</option><option value='12'>12개</option><option value='15'>15개</option>");
			<% elseif contentsEa=6 then %>
				$("#contentsEA").html("<option value='3'>3개</option><option value='6' selected>6개</option><option value='9'>9개</option><option value='12'>12개</option><option value='15'>15개</option>");
			<% elseif contentsEa=9 then %>
				$("#contentsEA").html("<option value='3'>3개</option><option value='6'>6개</option><option value='9' selected>9개</option><option value='12'>12개</option><option value='15'>15개</option>");
			<% elseif contentsEa=12 then %>
				$("#contentsEA").html("<option value='3'>3개</option><option value='6'>6개</option><option value='9'>9개</option><option value='12' selected>12개</option><option value='15'>15개</option>");
			<% elseif contentsEa=15 then %>
				$("#contentsEA").html("<option value='3'>3개</option><option value='6'>6개</option><option value='9'>9개</option><option value='12'>12개</option><option value='15' selected>15개</option>");
			<% end if %>
		<% elseif contentsCode="30" or contentsCode="31" then %>
			<% if contentsEa=1 then %>
				$("#contentsEA").html("<option value='1' selected>1개</option><option value='3'>3개</option>");
			<% else %>
				$("#contentsEA").html("<option value='1'>1개</option><option value='3' selected>3개</option>");
			<% end if %>
		<% end if %>
	});
//-->
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>메일진 템플릿 등록</h1>
	</div>
	<form name="frmReg" method="post" action="procTemplate.asp" onSubmit="return jsRegCode();">	
	<input type="hidden" name="mode" value="<%=sMode%>">
	<input type="hidden" name="idx" value="<%=idx%>">
	<div class="popContV19">
		<table class="tableV19A" id="table">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>메일진 종류</th>
					<td>
						<select name="mailzineKind" class="formControl">
							<option value="">선택</option>
							<% sbMailzineKindType (mailzineKind)%>
						</select>
					</td>
				</tr>
				<tr>
					<th>컨텐츠 종류</th>
					<td>
						<select name="contentsKind" class="formControl" onChange="fnContentsEaSet(this.value);">
							<option value="">선택</option>
							<% sbContentsKindType (contentsCode)%>
						</select>
					</td>
				</tr>
				<tr>
					<th>컨텐츠 수량</th>
					<td>
						<select name="contentsEA" id="contentsEA" class="formControl">
							<option value="">선택</option>
						</select>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1">저장</button>
	</div>
	</form>
	<div class="popHeadV19">
		<h1>메일진 템플릿 수정</h1>
	</div>
	<form name="frmSearch" method="post" action="popManageTemplate.asp">
	<input type="hidden" name="mode" value="S">
	<div class="popContV19">
		<div>
			<select name="mailzineKind" class="formControl" onChange="jsSearch(this.value);">
				<option value="">선택</option>
				<% sbMailzineKindType (mailzineKind)%>
			</select>
		</div>
        <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
            <%If isArray(arrList) THEN%>
            <h3 class="fs15">코드 리스트</h3>
            <table class="tableV19A tableV19B tMar10">
                <thead>
                    <tr>
                        <th></th>
						<th>코드값</th>
                        <th>코드명</th>
						<th>수량</th>
                        <th>정렬순서</th>
                        <th>처리</th>
                    </tr>
                <thead>
                <tbody id="subList">
				<%For intLoop = 0 To UBound(arrList,2)%>
                    <tr>
                        <td>
                            <span class="mdi mdi-equal cBl4 fs20"></span>
							<input type="hidden" name="idx" value="<%=arrList(0,intLoop)%>">
							<input type="hidden" name="viewidx" value="<%=arrList(4,intLoop)%>">
                        </td>
						<td><%=arrList(1,intLoop)%></td>
						<td><%=arrList(2,intLoop)%></td>
						<td><%=arrList(3,intLoop)%></td>
						<td><%=arrList(4,intLoop)%></td>
						<td><button class="btn4 btnGrey1" onClick="javascript:jsSetCode('<%=arrList(1,intLoop)%>','<%=mailzineKind%>','<%=arrList(0,intLoop)%>');return false;">수정</button>&nbsp;<button class="btn4 btnGrey1" onClick="javascript:jsDeleteCode('<%=arrList(0,intLoop)%>');return false;">삭제</button></td>
					</tr>
				<%Next%>
                </tbody>
			<%End if%>
		</div>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnBlue1" onClick="jsCodeSort(this.form);return false;">순서저장</button>
	</div>
	</form>
</div>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->