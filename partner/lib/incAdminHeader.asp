<!-- #include virtual="/lib/classes/adminpartner/partnerCls.asp" -->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
'###########################################################
' Description : 헤더
' Hieditor : 2016.11.24 정윤정 생성
'			 2016.12.29 한용민 수정(바코드출력 추가)
'###########################################################
%>
<%
if (session("ssGroupid")="") then
	Call Alert_move("세션이 종료되었습니다. \n재로그인후 사용하실수 있습니다.","/partner/index.asp")
	response.end
end if

'== 브랜드 리스트 ==========================
Dim HBList, conHB

'헤더_브랜드 리스트
Function fnGetHeaderBrand
	Dim ClsHPartner, arrHB
	set ClsHPartner = new CPartner
		ClsHPartner.FRectGroupID = session("ssGroupid")
		arrHB = ClsHPartner.fnGetBrandList
	set ClsHPartner = nothing
	fnGetHeaderBrand = arrHB
End Function

 if session("chkHBrand") <> 1 then
 	session("HBrandList") = fnGetHeaderBrand
 	session("chkHBrand") = 1
 end if
	HBList = session("HBrandList")
'== //브랜드 리스트 ==========================

'계약서 확인여부
dim noFinCtrExists, isNewContractTypeExists
dim NoConfirmPreContractID : NoConfirmPreContractID=-1
noFinCtrExists = isNotFinishNewContractExists(session("ssBctID"), session("ssGroupid"), isNewContractTypeExists)
if (Not noFinCtrExists) and (Not isNewContractTypeExists) then
    NoConfirmPreContractID = getLastPrecontractID(session("ssBctID"))
end if
%>
		<script type="text/javascript">
			//동일 업체의 다른브랜드로 로그인 변경
			function jsShiftBrand(sObj){
			 top.location.href="/partner/lib/shiftbrandHeader.asp?shiftid="+sObj.value;
			}

			//정보관리
			function jsPartnerInfo(){
				var popwin = window.open('/partner/company/company_Info_pop.asp' ,'op1','width=640,height=600,scrollbars=yes,resizable=yes');
			popwin.focus();
			}

			//파트별 담당자
			function jsManagerInfo(){
				var popwin = window.open('/common/pop_10x10_person.asp','op2','width=740,height=650,scrollbars=yes,resizable=no');
			popwin.focus();
			}

		//	//업체계약서 다운로드
		//	function jsContractInfo(ContractID){
		//	var url = "";
		//		if (ContractID < 1) {
		//			url = "/partner/company/contract/brand_contract_list_pop.asp";
		//	    }else{
		//	    	url ="/partner/company/contract/brand_contract_content_pop.asp"
		//	    }
		//	        var popwin = window.open(url+'?ContractID=' + ContractID,'popContract','width=900,height=700,scrollbars=yes,resizable=yes')
		//	        popwin.focus();
		//
		//	}

			//텐바이텐 약도
			function jstenbytenMap(){
				var popwin = window.open('/partner/common/map_10x10_pop.asp','op3','width=650,height=800,scrollbars=yes,resizable=yes')
				popwin.focus();
			}
		</script>
		<div class="headerWrap">
			<header>
				<div>
					<h1><a href="/">10x10 Partner</a></h1>
					<ul class="unb">
						<li class="bgNone"><span><strong><%=session("ssBctID")%>(<%=session("ssBctCname")%>)</strong>님, 안녕하세요!</span></li>
						<li class="bgNone">
							<%IF isArray(HBList) THEN%>
							<select name="selHB" title="브랜드명을 선택하세요" class="" onChange="jsShiftBrand(this)">
								 <% For conHB = 0 TO UBound(HBList,2)  %>
									<option value="<%=HBList(0,conHB)%>" <%IF (LCase(HBList(0,conHB))=LCase(session("ssBctId"))) then%>selected<%END IF%>><%=HBList(0,conHB)%>(<%= db2html(HBList(2,conHB)) %> <%= CHKIIF(HBList(3,conHB)="14","-아카데미","") %>)</option>
								 <% Next%>
							</select>
							<%END IF%>
						</li>
						<!--li class="bgNone"><a href="javascript:jsPartnerInfo();"><span>정보관리</span></a></li-->
						<li class="bgNone"><a href="javascript:jsManagerInfo();"><span>파트별 담당자</span></a></li>
						<!--li><a href="javascript:jsContractInfo('<%= NoConfirmPreContractID %>');"><span>업체계약서 다운로드</span></a></li-->
						<li><a href="<%=wwwURL%>/1463919" target="_blank"><span>소모품신청</span></a></li>
						<li><a href="javascript:jstenbytenMap();"><span>텐바이텐 약도</span></a></li>
						<li><a href="#" onclick="printbarcode_on_off_multi_upche(); return false;" ><span>바코드출력</span></a></li>
						<li class="bgNone"><a href="/login/dologout.asp" target="_top"><input type="button" value="로그아웃" class="btn cBk1" /></a></li>
					</ul>
				</div>
			</header>
		</div>
