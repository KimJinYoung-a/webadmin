<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : [CS]고객센터>>[CS]미처리CS리스트
' History : 이상구 생성
'			2023.11.15 한용민 수정(사용안하는 구업체어드민 폴더에서 cs폴더로 복사 이관)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_mifinishcls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim csdetailidx, ocsmifinishmaster, PreDispMail, isChulgoState, ioneas
    csdetailidx = requestCheckVar(request("csdetailidx"),10)

set ocsmifinishmaster = new CCSMifinishMaster
    ocsmifinishmaster.FRectCSDetailIDx = csdetailidx
    ocsmifinishmaster.getOneMifinishItem

	if ocsmifinishmaster.FtotalCount < 1 then
		ocsmifinishmaster.FRectCSDetailIDx = csdetailidx
		ocsmifinishmaster.FRectorder6MonthBefore = "Y"
		ocsmifinishmaster.getOneMifinishItem
	end if

if (ocsmifinishmaster.FResultCount<1) then
    response.write "검색결과가 없습니다."
    dbget.close() : response.end
end if

PreDispMail = (ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed) and (ocsmifinishmaster.FOneItem.FMifinishReason<>"05")
isChulgoState = (ocsmifinishmaster.FOneItem.Fdivcd = "A000") or (ocsmifinishmaster.FOneItem.Fdivcd = "A100")

set ioneas = new CCSASList
    ioneas.FRectCsAsID = ocsmifinishmaster.FOneItem.Fasid
    ioneas.GetOneCSASMaster

%>
<style type="text/css" >
.sale11px01 {font-family: dotum; FONT-SIZE: 11px; font-weight:bold ; COLOR: #b70606;}
</style>
<script language='javascript'>
//function getOnload(){
//    popupResize(700);
//}
//window.onload = getOnload;

function ShowDateBox(comp){
    var frm = comp.form;

    var iid = comp.id;
    var idiv = document.all.divipgodate;
    var isold = document.all.itemSoldOutFlag

	if (comp.value == "05") {
		// 품절출고불가
		idiv.style.display = "none";
		isold.style.display = "inline";
	} else {
		idiv.style.display = "inline";
		isold.style.display = "none";
	}
}

function ipgodateChange(comp){
    var v = comp.value;
    if (v.length<10) v = "YYYY-MM-DD";

    ShowDateBox(frmMisend.MifinishReason);
}

function MiFinishInput(){
    var frm = document.frmMisend;
    var today= new Date();
    today = new Date(today.getYear(),today.getMonth(),today.getDate());  //오늘도 가능하도록

    var inputdate;

    if (frm.MifinishReason.value.length<1){
        alert('미처리 사유를 입력하세요.');
        frm.MifinishReason.focus();
        return;
    }

    if (frm.MifinishReason.value == "05") {

    } else {
        var ipgodate = eval("frm.ipgodate");
        if (ipgodate.value.length!=10){
            alert('처리 예정일을 입력하세요.(YYYY-MM-DD)');
            ipgodate.focus();
            return;
        }

        inputdate = new Date(ipgodate.value.substr(0,4),ipgodate.value.substr(5,2)*1-1,ipgodate.value.substr(8,2));
        if (today>inputdate){
            alert('처리 예정일은 오늘 이후날짜로 설정이 가능합니다.');
            ipgodate.focus();
            return;
        }
    }

    if (confirm('미처리 사유를 저장 하시겠습니까?')){
	    frm.action = "/cscenter/mifinish/popMifinishInput_process.asp";
	    frm.submit();
	}
}

function getDiffDay(d1,d2){   // 두 날짜의 차이구함

  var v1=d1.split("-");
  var v2=d2.split("-");

  var a1=new Date(v1[0],v1[1],v1[2]);
  var a2=new Date(v2[0],v2[1],v2[2]);
  return parseInt((a2-a1)/(1000*3600*24));  //1000*3600*24 는 날의차이 만약 월의 차이를 구하고 싶다면 *30곱하면 월 12를 곱하면 년

}

</script>

<% if ocsmifinishmaster.FResultCount>0 then %>
<form name="frmMisend" method="post" action="/cscenter/mifinish/popMifinishInput_process.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="MiFinishInputOne">
<input type="hidden" name="csdetailidx" value="<%= ocsmifinishmaster.FOneItem.Fcsdetailidx %>">
<input type="hidden" name="asid" value="<%= ocsmifinishmaster.FOneItem.Fasid %>">
<input type="hidden" name="Sitemid" value="<%= ocsmifinishmaster.FOneItem.FItemID %>">
<input type="hidden" name="Sitemoption" value="<%= ocsmifinishmaster.FOneItem.FItemOption %>">
<input type="hidden" name="ischulgostate" value="<% if isChulgoState then %>Y<% end if %>">
<table width="610" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2">
	    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>CS미처리사유 입력</b>
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
    	<td width="130">구분</td>
    	<td width="480">
    		<font color="<%= ocsmifinishmaster.FOneItem.getDivcdColor %>"><%= ocsmifinishmaster.FOneItem.getDivcdStr %></font>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF" height="25">
    	<td width="130">상품코드</td>
    	<td width="480"><%= ocsmifinishmaster.FOneItem.FItemID %>

    	    <% if (ocsmifinishmaster.FOneItem.Fdeleteyn<>"N") then %>
				<b><font color="#CC3333">[취소CS]</font></b>
				<script language='javascript'>alert('취소된 CS 입니다.');</script>
			<% else %>
			    [정상CS]
			<% end if %>
    	</td>
    </tr>
	<tr bgcolor="#FFFFFF">
	    <td>이미지</td>
	    <td><img src="<%= ocsmifinishmaster.FOneItem.Fsmallimage %>" width="50" height="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>상품명</td>
	    <td><%= ocsmifinishmaster.FOneItem.FItemName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>옵션</td>
	    <td><%= ocsmifinishmaster.FOneItem.FItemoptionName %></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>접수수량</td>
	    <td><%= ocsmifinishmaster.FOneItem.FRegItemNo %>개
	    <% if (isChulgoState = True) then %>
		    <% if ( C_ADMIN_USER) then %>
		    (부족수량 <%= ocsmifinishmaster.FOneItem.Fitemlackno %>)
		    <% end if %>
		<% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>CS제목</td>
	    <td>
	    	<%= ioneas.FOneItem.FTitle %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>접수사유</td>
	    <td>
	    	<%= ioneas.FOneItem.Fgubun01Name %>&gt;&gt;<%= ioneas.FOneItem.Fgubun02Name %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>접수내용</td>
	    <td>
	    	<%= replace(ioneas.FOneItem.Fcontents_jupsu,VbCrlf,"<br>") %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>미처리사유</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	        	<%= ocsmifinishmaster.FOneItem.getMiFinishCodeName %>
	        <% else %>
	        <select name="MifinishReason" id="MifinishReason" class="select" onChange="ShowDateBox(this);">
				<option value=""></option>
				<% if (isChulgoState = True) then %>
					<option value="03" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="03","selected"," ") %> >출고지연</option>
					<option value="05" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="05","selected"," ") %> >품절출고불가</option>
					<option value="02" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="02","selected"," ") %> >주문제작(수입)</option>
					<option value="04" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="04","selected"," ") %> >예약배송</option>
					<option value="07" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="07","selected"," ") %> >고객지정배송</option>
				<% else %>
					<% if (C_ADMIN_USER) then %>
						<option value="25" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="25","selected"," ") %> >송장입력 안내</option>
						<option value="26" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="26","selected"," ") %> >반품불가 안내</option>
						<option value="21" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="21","selected"," ") %> >고객 부재</option>
						<option value="22" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="22","selected"," ") %> >고객 반품예정</option>
						<option value="23" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="23","selected"," ") %> >CS택배접수</option>
						<option value="12" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="12","selected"," ") %> >업체지연</option>
						<option value="41" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="41","selected"," ") %> >택배사 수거지연</option>
					<% else %>
						<!--
						<option value="11" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="11","selected"," ") %> >상품 회수이전</option>
						<option value="13" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="13","selected"," ") %> >삭제요청(고객 오입력)</option>
						<option value="14" <%= ChkIIF(ocsmifinishmaster.FOneItem.FMifinishReason="14","selected"," ") %> >기타</option>
						-->
					<% end if %>
				<% end if %>
			</select>
			<% end if %>
			<span id="itemSoldOutFlag" name="itemSoldOutFlag" style="display=none" align="right" >
			<input type="radio" name="itemSoldOut" value="N" checked >상품 품절처리
			<input type="radio" name="itemSoldOut" value="S">상품 일시품절처리
			</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
	    <td>처리예정일</td>
	    <td>
	        <% if (Not C_ADMIN_USER) and ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	        	<%= ocsmifinishmaster.FOneItem.FMifinishipgodate %>
	        <% else %>
		        <div id="divipgodate" name="divipgodate" <%= ChkIIF((ocsmifinishmaster.FOneItem.FMifinishReason <> "05" and Not IsNull(ocsmifinishmaster.FOneItem.FMifinishReason)),"style='display:inline'","style='display:none'") %> >
				    <input class="text" type="text" name="ipgodate" value="<%= ocsmifinishmaster.FOneItem.FMifinishipgodate %>" size="10" maxlength="10" onKeyup="ipgodateChange(this);">
				    <a href="javascript:calendarOpen(frmMisend.ipgodate);ipgodateChange(frmMisend.ipgodate);"><img src="/images/calicon.gif" border="0" align="top" height=20></a>
				</div>
			<% end if %>
	    </td>
	</tr>

	<% if (not C_ADMIN_USER) then %>
		<tr bgcolor="#FFFFFF" height="25">
		    <td>상세사유</td>
		    <td>
		    	<textarea class="textarea" name="finishmemo" cols="60" rows="6" class="a"><%= ioneas.FOneItem.Fcontents_finish %></textarea>
		    </td>
		</tr>
	<% end if %>

	<tr bgcolor="#FFFFFF" height="25">
	    <td>고객안내여부</td>
	    <td>
	    	<% if isChulgoState then %>
		        <% if (C_ADMIN_USER) then %>
		            <% if (ocsmifinishmaster.FOneItem.FisSendSms="Y") then %>
		                SMS발송완료/
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendSMS" type="checkbox" disabled  >SMS발송&nbsp;
		                <% else %>
		                <input name="ckSendSMS" type="checkbox" checked  >SMS발송&nbsp;
		                <% end if %>
		            <% else %>
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendSMS" type="checkbox" disabled  >SMS발송&nbsp;
		                <% else %>
		                <input name="ckSendSMS" type="checkbox" checked  >SMS발송&nbsp;
		                <% end if %>
		            <% end if %>

		            <% if (ocsmifinishmaster.FOneItem.FisSendEmail="Y") then %>
		                MAIL발송완료/
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendEmail" type="checkbox" disabled  >MAIL발송&nbsp;
		                <% else %>
		                <input name="ckSendEmail" type="checkbox" checked  >MAIL발송&nbsp;
		                <% end if %>
		            <% else %>
		                <% if (ocsmifinishmaster.FOneItem.FMifinishReason="05") then %>
		                <input name="ckSendEmail" type="checkbox" disabled  >MAIL발송&nbsp;
		                <% else %>
		                <input name="ckSendEmail" type="checkbox" checked  >MAIL발송&nbsp;
		                <% end if %>
		            <% end if %>
		        <% else %>
	    	        <% if ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
	    	            <!-- 고객안내가 완료된 건은 미출고사유 및 출고예정일 수정 불가 -->
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendSms="Y","SMS발송완료/","") %>
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendEmail="Y","MAIL발송완료/","") %>
	    	            <%= CHKIIF(ocsmifinishmaster.FOneItem.FisSendCall="Y","통화안내완료","") %>
	    	        <% else %>
	        	        <input name="ckSendSMS" type="checkbox" checked disabled >SMS발송
	        	        &nbsp;
	        	        <input name="ckSendEmail" type="checkbox" checked disabled >MAIL발송
	    	        <% end if %>
	    	    <% end if %>
			<% else %>
    	        <input name="ckSendSMS" type="checkbox" disabled >SMS발송
    	        &nbsp;
    	        <input name="ckSendEmail" type="checkbox" disabled >MAIL발송
			<% end if %>
	    </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
	    <td colspan="2">
	    	<font color="blue">
	    	<% if isChulgoState then %>
		    	미출고 사유가 출고지연 및 주문제작(수입)일 경우, 아래의 내용으로 고객님께 SMS와 메일이 발송됩니다.<br>
		    	고객님께 안내된 출고예정일을 꼭 지켜주시기 바라며, 변동사항이 생길경우, 고객센터로 연락 부탁드립니다.<br>
		    	</font>
		    	<font color="red">
		       	품절출고불가인 경우, 고객님께 SMS 및 메일이 발송되지 않으며, 텐바이텐고객센터에서<br>
		    	별도로 고객님께 연락을 드릴 예정입니다.
		    	</font>
		    <% else %>

		<% end if %>
	    </td>
	</tr>
	<tr height="20" bgcolor="<%= adminColor("tabletop") %>">
	    <td colspan="2" align="center">
	    <% if (C_ADMIN_USER) then %>
	        <% if (ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed) and (ocsmifinishmaster.FOneItem.FisSendSms="Y") and (ocsmifinishmaster.FOneItem.FisSendEmail="Y") then %>
    	    기존 저장된 내역입니다.<br>
    	    <input type="button" class="button" value="미처리 사유 다시 저장" onclick="MiFinishInput();">
    	    <% else %>
	        <input type="button" class="button" value="미처리 사유 저장" onclick="MiFinishInput();">
	        <% end if %>
	    <% else %>
    	    <% if ocsmifinishmaster.FOneItem.isMifinishAlreadyInputed then %>
    	    수정 불가
    	    <% else %>
    	    <input type="button" class="button" value="미처리 사유 저장" onclick="MiFinishInput();">
    	    <% end if %>
    	<% end if %>
	    </td>
	</tr>
</table>
</form>
<br>
<% else %>
<table width="600">
<tr>
    <td align="center">취소된 CS이거나 해당 CS 내역이 없습니다.</td>
</tr>
</table>
<% end if %>

<%
set ocsmifinishmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
