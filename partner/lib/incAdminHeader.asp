<!-- #include virtual="/lib/classes/adminpartner/partnerCls.asp" -->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
'###########################################################
' Description : ���
' Hieditor : 2016.11.24 ������ ����
'			 2016.12.29 �ѿ�� ����(���ڵ���� �߰�)
'###########################################################
%>
<%
if (session("ssGroupid")="") then
	Call Alert_move("������ ����Ǿ����ϴ�. \n��α����� ����ϽǼ� �ֽ��ϴ�.","/partner/index.asp")
	response.end
end if

'== �귣�� ����Ʈ ==========================
Dim HBList, conHB

'���_�귣�� ����Ʈ
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
'== //�귣�� ����Ʈ ==========================

'��༭ Ȯ�ο���
dim noFinCtrExists, isNewContractTypeExists
dim NoConfirmPreContractID : NoConfirmPreContractID=-1
noFinCtrExists = isNotFinishNewContractExists(session("ssBctID"), session("ssGroupid"), isNewContractTypeExists)
if (Not noFinCtrExists) and (Not isNewContractTypeExists) then
    NoConfirmPreContractID = getLastPrecontractID(session("ssBctID"))
end if
%>
		<script type="text/javascript">
			//���� ��ü�� �ٸ��귣��� �α��� ����
			function jsShiftBrand(sObj){
			 top.location.href="/partner/lib/shiftbrandHeader.asp?shiftid="+sObj.value;
			}

			//��������
			function jsPartnerInfo(){
				var popwin = window.open('/partner/company/company_Info_pop.asp' ,'op1','width=640,height=600,scrollbars=yes,resizable=yes');
			popwin.focus();
			}

			//��Ʈ�� �����
			function jsManagerInfo(){
				var popwin = window.open('/common/pop_10x10_person.asp','op2','width=740,height=650,scrollbars=yes,resizable=no');
			popwin.focus();
			}

		//	//��ü��༭ �ٿ�ε�
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

			//�ٹ����� �൵
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
						<li class="bgNone"><span><strong><%=session("ssBctID")%>(<%=session("ssBctCname")%>)</strong>��, �ȳ��ϼ���!</span></li>
						<li class="bgNone">
							<%IF isArray(HBList) THEN%>
							<select name="selHB" title="�귣����� �����ϼ���" class="" onChange="jsShiftBrand(this)">
								 <% For conHB = 0 TO UBound(HBList,2)  %>
									<option value="<%=HBList(0,conHB)%>" <%IF (LCase(HBList(0,conHB))=LCase(session("ssBctId"))) then%>selected<%END IF%>><%=HBList(0,conHB)%>(<%= db2html(HBList(2,conHB)) %> <%= CHKIIF(HBList(3,conHB)="14","-��ī����","") %>)</option>
								 <% Next%>
							</select>
							<%END IF%>
						</li>
						<!--li class="bgNone"><a href="javascript:jsPartnerInfo();"><span>��������</span></a></li-->
						<li class="bgNone"><a href="javascript:jsManagerInfo();"><span>��Ʈ�� �����</span></a></li>
						<!--li><a href="javascript:jsContractInfo('<%= NoConfirmPreContractID %>');"><span>��ü��༭ �ٿ�ε�</span></a></li-->
						<li><a href="<%=wwwURL%>/1463919" target="_blank"><span>�Ҹ�ǰ��û</span></a></li>
						<li><a href="javascript:jstenbytenMap();"><span>�ٹ����� �൵</span></a></li>
						<li><a href="#" onclick="printbarcode_on_off_multi_upche(); return false;" ><span>���ڵ����</span></a></li>
						<li class="bgNone"><a href="/login/dologout.asp" target="_top"><input type="button" value="�α׾ƿ�" class="btn cBk1" /></a></li>
					</ul>
				</div>
			</header>
		</div>
