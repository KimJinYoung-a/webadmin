<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 브랜드 계약 관리
' Hieditor : 2009.04.07 서동석 생성
'			 2010.05.26 한용민 수정
'###########################################################
%>
<!-- #include virtual="/lib/CheckLoginReDirect.asp" -->
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim ContractID, makerid , i , sqlStr , ocontract, ocontractList ,opartner , onoffgubun
	ContractID  = requestCheckVar(request("ContractID"),100)
	makerid     = session("ssBctID")

set ocontractList = new CPartnerContract
ocontractList.FRectMakerid = makerid
if makerid<>"" then
    ocontractList.GetMakerValidContractList
end if

if (ContractID="") or (ContractID="0") then
    if (ocontractList.FResultCount>0) then
        ContractID = ocontractList.FItemList(0).FContractID
    end if
end if


set ocontract = new CPartnerContract
ocontract.FRectContractID = ContractID
ocontract.FRectMakerid = makerid
if ContractID<>"" then
    ocontract.getOneContract
end if

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if (makerid<>"") then
    opartner.GetOnePartnerNUser
end if

if ocontract.FResultCount>0 then
	if ocontract.FOneItem.FContractType <> "" then
		sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
		sqlStr = sqlStr & " where contractType=" & ocontract.FOneItem.FContractType
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
		    onoffgubun = rsget("onoffgubun")
		end if
		rsget.Close
	end if
end if
%>

<script language='javascript'>

//window.resizeTo(600,600);

function changeContract(comp){
    document.frmResearch.ContractID.value = comp.value;
    document.frmResearch.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="right">
	        <!-- select Box -->
	        <select class="select" name="ContractID" onChange="changeContract(this);">
		        <% for i=0 to ocontractList.FResultCount-1 %>
		        <option value="<%= ocontractList.FItemList(i).FContractID %>" <% if CStr(ocontractList.FItemList(i).FContractID)=ContractID  then response.write "selected" %> >[<%= ocontractList.FItemList(i).FContractNo %>] <%= ocontractList.FItemList(i).FContractName %>
		        <% next %>
	        </select>
	    </td>
	</tr>
	<% if ocontract.FResultCount>0 then %>
	<tr bgcolor="#FFFFFF">
	    <td>
	        <table width="100%" border="0" cellspacing="1" cellpadding="1" class="a" >
	        <tr>
	            <td>
	            안녕하세요<br>
	            (주)텐바이텐과 좋은 인연으로 만나게 되어 반갑습니다.<br>
	            <br>
	            아래와 같이 계약이 진행되오니 <br>
	            계약서 진행사항을 꼼꼼히 읽어주신 후 <br>
	            일정에 맞추어 계약서를 우편으로 발송해 주시면 감사하겠습니다.<br>
	            </td>
	        </tr>
	        <tr>
	            <td><br>계약서 명 :  <%= ocontract.FOneItem.FContractName %>  </td>
	        <tr>
	        <tr>
	            <td>계약서 번호 : <%= ocontract.FOneItem.FcontractNo %>  </td>
	        <tr>
	        <% if ocontract.FOneItem.FContractState>=7 then %>
	        <tr>
	            <td>상태 : 계약완료 (완료일 : <%= ocontract.FOneItem.FFinishDate %>) </td>
	        <tr>
	        <% else %>
	        <tr>
	            <td>
	            <% if onoffgubun = "ON" then %>	
		            ▶ 계약서 진행대상 : 온라인 입점 전 브랜드  <br>
		             - 오프라인에만 진행하는 브랜드는 대상에서 제외됩니다. <br>
		             (오프라인 입점시에는 오프라인 담당자가 개별적으로 보내드립니다.)
		        <% else %>
		            ▶ 계약서 진행대상 : 오프라인 입점 전 브랜드  <br>		        
		        <% end if %>     
	            <br>
	            <br>
	            
	            ▶ 계약서 다운방법 <br>
	            - 아래 [계약서 다운로드] 클릭하여 다운받은 후 내용확인 및 기재사항을 기재해주세요!!<br>
	            - 계약서 날인 하는 방법은 [계약flow 다운로드] 다운받으신 후 그 방법대로 해주시기 바랍니다.<br>
	
	            <a href="/designer/company/downLoadContract.asp?ContractID=<%= ContractID %>" target="iTargetFrm"><b><font color="blue">[계약서 다운로드]</font></b></a>
	            <br>
	            <a href="/designer/company/contractflow.ppt" target="_blank"><b><font color="blue">[계약flow 다운로드]</font></b></a>
	            
	            <br><br>
	            ▶ 필수 확인사항 (반드시 두번 세번 확인해주세요)  <br>
	            수수료, 결제일 이 두가지는 맞는 지 꼭 확인해주셔야 합니다. 
	            <br><br>
	            ▶ 업체측에서의 필수 기재사항 (꼭!! 직접 기재하셔야 할 부분)<br>
	            - 표지(첫장)의 계약담당자 기재 : 협력업체의 대표이사 또는 계약을 실제로 진행하시는 담당자 성함<br>
	            <% if ocontract.FOneItem.FContractType=5 then %>
	            - 배송책임자 성함<br>
	            <% end if %>
	            - 마지막 장의 '을'의 대표이사 주민등록번호 및 주소 기재 : 사업자등록증의 대표자 주민번호 및 주소여야 합니다.<br>
	            - 법인사업자의 경우 마지막장의 대표이사 주민번호 및 주소를 생략하셔도 되며, '갑'사업자 정보 '을' 사업자 정보만 있으면 됩니다. 
	
	            <br><br>
	            ▶ 진행절차 :  <br>
	            ① 계약서 다운로드 <br>
	            ② 협력업체에서 계약서 확인 후 날인 / 2부 우편발송  <br>
	            ③ 텐바이텐에서 계약서 우편 수령확인 <br>
	            ④ 텐바이텐에서 협력업체로 계약서 1부 발송 / 계약완료 
	            <br>
	            <br>
	            
	            ▶ 계약서 보내시는 곳 <br>
	            <% if onoffgubun = "ON" then %>
		            주소 : 서울시 종로구 동숭동 1-45번지 자유빌딩 5층 텐바이텐 <br>
		            담당자 : <%= ocontract.FoneItem.Fusername %> <br>
		            tel : <%= ocontract.FoneItem.Finterphoneno %> (내선 <%= ocontract.FoneItem.Fextension %>) / 직통 : <%= ocontract.FoneItem.Fdirect070 %><br>
		            fax : 02-2179-9244 <br>
	            <% else %>
		            주소 : 서울시 종로구 동숭동 1-74 에버리치 홀딩스빌딩 6층 텐바이텐 오프라인 사무실<br>
		            담당자 : <%= ocontract.FoneItem.Fusername %> <br>
		            tel : <%= ocontract.FoneItem.Finterphoneno %> (내선 <%= ocontract.FoneItem.Fextension %>) / 직통 : <%= ocontract.FoneItem.Fdirect070 %><br>
		            fax : 02-2179-9058 <br>
	            <% end if %>
	            
	            <br>
	            
	            
	            <% if onoffgubun = "ON" then %>
	            <!-- 디비화 필요..
	            ◈ 가구패브릭 / 조명데코 : 이윤선 대리 (내선 153 ) <br>
                ◈ 우먼 맨 패션 : 김지웅 대리 (내선 154 )<br>
                ◈ 디자인문구 /오피스개인 :오영섭 주임 (내선 152)<br>
                ◈ 카메라, book, baby : 최맑은소리 주임 (내선159)<br>
                ◈ 키덜트 취미 : 신은영 (내선 157)<br>
                ◈ 주방욕실 음악 테이스트 : 조영인 (내선 156)<br>
                ◈ 패션잡화 쥬얼리 : 문주희 (내선 155)<br>
                -->
	            <!-- 문구사무:오영섭 주임 (내선 152) / 리빙주방:이윤선 대리 (내선 153 ) <br>
	            패션쥬얼리:김지웅 대리 (내선 154 ) / 키덜트: 신은영 (내선 157) <br> -->
				<% end if %>	            
	            
	            <br><br>
	            ▶ 우편발송시 함께 보내셔야 할 서류 <br>
	            - 날인된 계약서 2부 <br>
	            - 결제통장 사본 <br>
	            - 사업자 등록증 사본 <br>
	            - 인감증명서 원본 (계약서에 날인한 도장) 
	            <br>
	            <br>
	            ▶ 기 타 <br>
	            텐바이텐 내에 진행하는 브랜드 아이디가 2개 이상일 경우 <br>
	            계약서는 각 브랜드 아이디마다 작성을 해주셔야 하며, <br>
	            관련서류(사업자등록증,인감증명서,결제통장)은 1부만 주셔도 됩니다. 
	            <br>
	            <br>
	            ▶ 계약서 내용 상의 궁금한 점은 각 담당MD에게 문의 하시기 바랍니다.
	            
	            <!--
	            ▶ 중요변경사항  <br>
	            - 사업자등록증 별로 계약 -> 브랜드별 계약  <br>
	            (예를 들어 한 사업자에서 3개의 브랜드를 운영할 경우 각 브랜드마다 계약하여 총 3개의 계약서를 날인해주셔야 합니다.) <br>
	            - 지적재산권 내용의 추가 및 상품등록 게시물의 규제 강화
	            <br><br>
	            -->
	            
	            <% if onoffgubun = "ON" then %>
	            <!--
	            ▶ 계약서 내용 상의 궁금한 점은 아래로 연락주세요 (단 수수료, 정산일 등은 담당 MD에게 문의)<br>
	            
	            서울텐바이텐 사업팀, 마케팅전략파트 <br>
	            대리 이설은<br>
	            TEL : 02-554-2033(#143)<br>
	            FAX : 02-2179-9245<br>
	            E.MAIL : snowsilver@10x10.co.kr<br>
	            WEB : www.10x10.co.kr<br>
	            -->
				<% end if %>
	            </td>
	        </tr>
	        <% end if %>
	        </table>
	    </td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
	    <td height="50" align="center">
	    [<%= makerid %> : 선택된 계약서가 없습니다. 먼저 계약서를 선택해 주세요.]
	    </td>
	</tr>
	<% end if %>
</table>

<form name="frmResearch">
<input type="hidden" name="ContractID" value="<%= ContractID %>">
</form>
<iframe name="iTargetFrm" id="iTargetFrm" src="" width="1" height="1" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<% 
set ocontract = Nothing
set ocontractList = Nothing
set opartner = Nothing
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->