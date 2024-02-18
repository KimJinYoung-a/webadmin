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
Dim selCodeType
Dim sCodeType,iCodeValue, sCodeDesc, iCodeSort, blnUsing, sCodeDispYN
Dim clsCode, sMode

iCodeValue  = requestCheckVar(Request("iCV"),10)	
selCodeType = requestCheckVar(Request("selCT"),20)
sCodeType   = requestCheckVar(Request("sCT"),20)
blnUsing = "Y"
sCodeDispYN ="Y"
sMode ="I"
if selCodeType="" then selCodeType="mailzineKind"

 Set clsCode = new CEventCommonCode  	
 	IF iCodeValue <> "" THEN
 		sMode ="U"
 		clsCode.FCodeType  = sCodeType 
 		clsCode.FCodeValue = iCodeValue
 		clsCode.fnGetEventCodeCont 		
 		sCodeDesc = clsCode.FCodeDesc
 		iCodeSort = clsCode.FCodeSort
 		blnUsing  = clsCode.FCodeUsing
 		sCodeDispYN=clsCode.FCodeDispYN
   	END IF
 		 
   	clsCode.FCodeType = selCodeType
   	arrList = clsCode.fnGetEventCodeList
 Set clsCode = nothing 
IF isnull(iCodeSort) or iCodeSort = "" THEN iCodeSort = 0 
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
	function jsSetCode(iCodeValue,selCodeType){	
		self.location.href = "PopManageCode.asp?iCV="+iCodeValue+"&sCT="+selCodeType+"&selCT="+selCodeType;
	}
	
	//코드 검색
	function jsSearch(){
		document.frmSearch.submit();
	}

	function jsCodeSort(){
		document.frmSearch.action="procCode.asp";
		document.frmSearch.submit();
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.iCV.value) {
			alert("코드값을 입력해 주세요");
			frm.iCV.focus();
			return false;
		}
			 
		if(!frm.sCD.value) {
			alert("코드명을 입력해 주세요");
			frm.sCD.focus();
			return false;
		}
		frm.submit();
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
	});

	function jsDeleteCode(iCodeValue,selCodeType){
		if(confirm("삭제하시겠습니까?")){
			document.frmReg.sM.value="D";
			document.frmReg.iCV.value=iCodeValue;
			document.frmReg.sCT.value=selCodeType;
			document.frmReg.submit();
		}
	}
//-->
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>메일진 코드 등록</h1>
	</div>
	<form name="frmReg" method="post" action="procCode.asp">	
	<input type="hidden" name="sM" value="<%=sMode%>">
	<div class="popContV19">
		<table class="tableV19A" id="table">
			<colgroup>
				<col style="width:150px;">
				<col style="width:auto;">
			</colgroup>
			<tbody>
				<tr>
					<th>코드타입</th>
					<td>
						<select name="sCT" class="formControl">						
						<% sbOptCodeType (sCodeType)%>					
						</select>
					</td>
				</tr>
				<tr>
					<th>코드값</th>
					<td>
						<%IF iCodeValue ="" THEN%>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="iCV" id="iCV" maxlength="10">
						<%ELSE%>
						<%=iCodeValue%><input type="hidden" size="4" maxlength="10" name="iCV" value="<%=iCodeValue%>">
						<%END IF%>
					</td>
				</tr>
				<tr>
					<th>코드명</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="sCD" id="sCD" maxlength="16" value="<%=sCodeDesc%>">
					</td>
				</tr>
				<tr>
					<th>코드 정렬순서</th>
					<td>
						<input type="text" class="formControl formControl550" placeholder="입력하세요." name="iCS" id="iCS" maxlength="10" value="<%=iCodeSort%>">
					</td>
				</tr>
				<tr>
					<th>코드 전시여부</th>
					<td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoD" value="Y"<% if sCodeDispYN="Y" then response.write " checked" %>>
                                전시
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoD" value="N"<% if sCodeDispYN="N" then response.write " checked" %>>
                                전시안함
                                <i class="inputHelper"></i>
                            </label>
                        </div>
					</td>
				</tr>
				<tr>
					<th>사용여부</th>
					<td>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoU" value="Y"<% if blnUsing="Y" then response.write " checked" %>>
                                사용
                                <i class="inputHelper"></i>
                            </label>
                        </div>
                        <div class="formInline">
                            <label class="formCheckLabel">
                                <input type="radio" class="formCheckInput" name="rdoU" value="N"<% if blnUsing="N" then response.write " checked" %>>
                                사용안함
                                <i class="inputHelper"></i>
                            </label>
                        </div>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	<div class="popBtnWrapV19">
		<button class="btn4 btnWhite1" onClick="self.close();">취소</button>
		<button class="btn4 btnBlue1" onClick="jsRegCode();return false;">저장</button>
	</div>
	</form>
	<div class="popHeadV19">
		<h1>메일진 코드 수정</h1>
	</div>
	<form name="frmSearch" method="post" action="PopManageCode.asp">
	<input type="hidden" name="sM" value="S">
	<div class="popContV19">
		<div>
			<select name="selCT" onChange="jsSearch();" class="formControl">						
			<% sbOptCodeType (selCodeType)%>					
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
                        <th>정렬순서</th>
                        <th>전시여부</th>
                        <th>사용여부</th>
                        <th>처리</th>
                    </tr>
                <thead>
                <tbody id="subList">
				<%For intLoop = 0 To UBound(arrList,2)%>
                    <tr>
                        <td>
                            <span class="mdi mdi-equal cBl4 fs20"></span>
							<input type="hidden" name="code_value" value="<%=arrList(1,intLoop)%>">
							<input type="hidden" name="viewidx" value="<%=arrList(4,intLoop)%>">
                        </td>
						<td><%=arrList(1,intLoop)%></td>
						<td><%=arrList(2,intLoop)%></td>
						<td><%=arrList(4,intLoop)%></td>
						<td><%=arrList(5,intLoop)%></td>
						<td><%=arrList(3,intLoop)%></td>
						<td><button class="btn4 btnGrey1" onClick="javascript:jsSetCode('<%=arrList(1,intLoop)%>','<%=arrList(0,intLoop)%>');return false;">수정</button><% if arrList(0,intLoop)="mailzineKind" then %>&nbsp;<button class="btn4 btnGrey1" onClick="javascript:jsDeleteCode('<%=arrList(1,intLoop)%>','<%=arrList(0,intLoop)%>');return false;">삭제</button><% end if %></td>
					</tr>
				<%Next%>
                </tbody>
            </table>
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