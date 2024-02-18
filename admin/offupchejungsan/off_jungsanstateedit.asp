<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/jungsan_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->

<%
dim idx
idx = requestCheckvar(request("idx"),10)


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FRectIdx = idx
''ooffjungsan.FRectMakerid = 업체일경우 session 업체아이디
ooffjungsan.GetOneOffJungsanMaster


if (ooffjungsan.FResultCount<1) then
    response.write "<script>alert('검색 결과가 없습니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

dim makerid
makerid = ooffjungsan.FOneItem.Fmakerid


dim opartner
set opartner = new CPartnerUser
opartner.FCurrpage = 1
opartner.FRectDesignerID = makerid
opartner.FPageSize = 1
opartner.GetOnePartnerNUser


dim isPreJungsanFixAvail    ''계산서 없이 정산 확정 가능여부 - 사입브랜드, 수시결제

isPreJungsanFixAvail = ((opartner.FoneItem.FpurchaseType="4" or opartner.FoneItem.FpurchaseType="5") and (opartner.FoneItem.Fjungsan_date_off="수시"))
%>
<script language='javascript'>
function changeStep(frm,mode){
    var taxregdate;
    if (mode=='step1to3'){
        //정산확정 진행
        //if (!(calendarOpen2(frm.taxregdate))) return;

        if (confirm('계산서 발행일 [' + frm.taxregdate.value + '] 진행 하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }else if (mode=='step1to3noTax'){
        //정산확정 선결제 가능 브랜드 진행
        if (confirm('계산서 없이 정산 확정 진행 하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }else if (mode=='step3to0'){
        //정산확정->수정중상태로 변경 : 제약조건 tax log에서 삭제 되어야 함.
        if (confirm('정산 확정중 수정상태로 진행하려면 \n전자세금계산서 발행 로그를 먼저 삭제 하셔야 \n 수정중 상태로 바뀔수 있습니다. \n 진행 하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }else if (mode=='step3to7'){
        //입금완료 진행
        //if (!(calendarOpen2(frm.ipkumdate))) return;

        if (confirm('입금일 [' + frm.ipkumdate.value + '] 진행 하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }else if (mode=='deltaxinfo'){
        if (confirm('계산서 정보를 삭제하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }else{
        if (confirm('수정 하시겠습니까?')){
            frmMaster.mode.value = mode;
            frmMaster.submit();
        }
    }
}

function SaveFrm(frm,mode){

    if (frm.taxtype.value.length<1){
        alert('과세구분을 선택하세요.');
        frm.taxtype.focus();
        return;
    }

    if (frm.ispreFixTaxDateForce.checked){
        if (frm.preFixedTaxDate.value.length!=10){
            alert('계산서 발행 지정일을 입력하세요(YYYY-MM-DD)');
            frm.preFixedTaxDate.focus();
            return;
        }
    }
    /*
    if (frm.isrefPay.checked){
        if (frm.refPayreqIdx.value.length<1){
            alert('결제요청서 번호를 입력하세요.');
            frm.refPayreqIdx.focus();
            return;
        }
    }
    */
    if (confirm('저장  하시겠습니까?')){
        frmMaster.mode.value = mode;
        frmMaster.submit();
    }
}

function saveGroupid(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editGroupid";
		frm.submit();
	}
}

function saveJacct(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editJAcctCd";
		frm.submit();
	}
}

function saveAvailNeo(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.mode.value="editAvailNeo";
		frm.submit();
	}
}

function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}


function jsGetTax(ibizNo, itotSum){
	var sSearchText = ibizNo;
	var itotSum = itotSum;
	var winTax = window.open("/admin/tax/popSetEseroTax.asp?sST="+sSearchText+"&totSum="+itotSum+"&tgType=NRM","popGetTaxInfo","width=1200, height=800, resizable=yes, scrollbars=yes");
	winTax.focus();
}

function fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP){
    var frm = document.frmMaster;
    frm.taxregdate.value = dID;
    frm.eseroEvalSeq.value = eTax;

    //발행업체 지정
    var mayApCd = eTax.substring(8,16);
    if (mayApCd=="10000000"){
        //국세청
        frm.billsiteCode.value = 'E';
    }else if(mayApCd=="10000966"){
        //빌365
        frm.billsiteCode.value = 'B';
    }else{
        //기타
        frm.billsiteCode.value = 'Y';
    }
}

function fillTaxInfoWithPayreqIdx(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP,prIdx){
    fillTaxInfo(eTax,iDK,iVK,dID,sInm,mTP,mSP,mVP);

    var frm = document.frmMaster;
    frm.refPayreqIdx.value=prIdx;
    if (frm.refPayreqIdx.value.length>0){
        frm.isrefPay.checked=true;
    }
}

function jsNewRegXML(){
    var winD = window.open("/admin/tax/popRegfileXML.asp","popDXML","width=600, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}


function jsNewRegHand(){
    var winD = window.open("/admin/tax/popRegfileHand.asp","popDHand","width=860, height=400, resizable=yes, scrollbars=yes");
	winD.focus();
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frm" method="get" action="">
    <input type="hidden" name="idx" value="<%= idx %>">
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top" bgcolor="F4F4F4" width="530">
            <%= ooffjungsan.FOneItem.FTitle %>&nbsp;<%= ooffjungsan.FOneItem.Fmakerid %>&nbsp;&nbsp;
            총 정산액 : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_jungsanprice,0) %>&nbsp;&nbsp;
            총 판매상품수 : <%= FormatNumber(ooffjungsan.FOneItem.Ftot_itemno,0) %>
        </td>
        <td valign="top" bgcolor="F4F4F4" align="right">
        &nbsp;
        <!--
            <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        -->
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" width="100">브랜드ID</td>
        <td><%= makerid %></td>
    </tr>
</table>

<br>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="off_jungsan_process.asp">
<input type="hidden" name="masteridx" value="<%= idx %>">
<input type="hidden" name="mode" value="">

    <tr bgcolor="#FFFFFF">
    	<td bgcolor="#E6E6E6" width="100">정산대상년월</td>
    	<td bgcolor="#FFFFFF"><%= ooffjungsan.FOneItem.FYYYYMM %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="#E6E6E6" width="100">정산방식구분</td>
    	<td bgcolor="#FFFFFF"><%= ooffjungsan.FOneItem.getJGubunName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td bgcolor="#E6E6E6" width="100">Title</td>
    	<td bgcolor="#FFFFFF"><%= ooffjungsan.FOneItem.FTitle %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >진행상태</td>
        <td>

            <font color="<%= ooffjungsan.FOneItem.GetStateColor %>"><%= ooffjungsan.FOneItem.GetStateName %></font>

            <% if ooffjungsan.FOneItem.Ffinishflag="0" then %>
                &nbsp;&nbsp;&nbsp;
                <input type="button" value="업체 확인중으로 변경" onClick="changeStep(frmMaster,'step0to1');">
            <% elseif (ooffjungsan.FOneItem.Ffinishflag="1") or (ooffjungsan.FOneItem.Ffinishflag="2") then %>
            <!-- 업체 확인중 상태 -->
                &nbsp;&nbsp;&nbsp;
                <input type="button" value="수정중 상태로 변경" onClick="changeStep(frmMaster,'step1to0');">
                &nbsp;
                <input type="button" value="정산 확정 변경(수기 계산서)" onClick="changeStep(frmMaster,'step1to3');">
                <% if (isPreJungsanFixAvail) then %>
                &nbsp;
                <input type="button" value="정산 확정 변경(선결제)" onClick="changeStep(frmMaster,'step1to3noTax');">
                <% end if %>
            <% elseif ooffjungsan.FOneItem.Ffinishflag="3" then %>
            <!-- 정산 확정 상태 -->
                &nbsp;&nbsp;&nbsp;
                <input type="button" value="수정중 상태로 변경" onClick="changeStep(frmMaster,'step3to0');">
                <input type="button" value="입금완료 상태로 변경" onClick="changeStep(frmMaster,'step3to7');">
            <% elseif ooffjungsan.FOneItem.Ffinishflag="7" then %>
            <!-- 입금완료상태 -->
                &nbsp;&nbsp;&nbsp;
                <% if (C_ADMIN_AUTH) then %>
                    <input type="button" value="수정중 상태로 변경" onClick="changeStep(frmMaster,'step7to0');">
                <% else %>
                    ( 입금완료 상태중 상태변경은  관리자에게 문의 하세요 )
                <% end if %>
            <% end if %>
        </td>
    </tr>

    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >과세구분</td>
        <td>
            <select name="taxtype" >
    		<option value="" <% if IsNULL(ooffjungsan.FOneItem.Ftaxtype) or (ooffjungsan.FOneItem.Ftaxtype="") then response.write "selected" %> >
    		<option value="01" <% if ooffjungsan.FOneItem.Ftaxtype="01" then response.write "selected" %> >과세
    		<option value="02" <% if ooffjungsan.FOneItem.Ftaxtype="02" then response.write "selected" %> >면세
    		<option value="03" <% if ooffjungsan.FOneItem.Ftaxtype="03" then response.write "selected" %> >원천
    		</select>

            <%= ooffjungsan.FOneItem.Ftaxtype %>
            (기본설정: <b><%= opartner.FOneItem.Fjungsan_gubun %></b>)
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >차수구분</td>
        <td>
            <%= ooffjungsan.FOneItem.Fdifferencekey %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >계산서발행일</td>
        <td>
            <input type="text" name="taxregdate" value="<%= ooffjungsan.FOneItem.Ftaxregdate %>" size="10" maxlength="10">
            <a href="javascript:calendarOpen(frmMaster.taxregdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
            <% if Not IsNULL(ooffjungsan.FOneItem.Ftaxinputdate) then %>
            (<%= ooffjungsan.FOneItem.Ftaxinputdate %>)
            <% end if %>

            <% If ISNULL(ooffjungsan.FOneItem.Ftaxlinkidx) then %>
          	&nbsp;
          	<input type="button" value="선택" onClick="jsGetTax('<%= REplace(ooffjungsan.FOneItem.Fcompany_no,"-","") %>','<%= ooffjungsan.FOneItem.Ftot_jungsanprice %>');">
          	<input type="button" value="XML" onClick="jsNewRegXML();">
          	<input type="button" value="종이계산서입력" onClick="jsNewRegHand();">
          	<% else %>
              	<% if (ooffjungsan.FOneItem.Ffinishflag="0" or ooffjungsan.FOneItem.Ffinishflag="1" or ooffjungsan.FOneItem.Ffinishflag="2") then %>
              	<input type="button" value="계산서발행정보삭제" onClick="changeStep(frmMaster,'deltaxinfo')">
              	<% end if %>
          	<% end if %>
            <br>
            <input type="hidden" name="taxlinkidx" value="<%= ooffjungsan.FOneItem.Ftaxlinkidx %>">
            <% if isNULL(ooffjungsan.FOneItem.Ftaxlinkidx) then %>
                <% call DrawBillSiteCombo("billsiteCode",ooffjungsan.FOneItem.FbillsiteCode) %>
            <% else %>
                <input type="hidden" name="billsiteCode" value="<%= ooffjungsan.FOneItem.FbillsiteCode %>">
                <%= ooffjungsan.FOneItem.FbillSiteName %>
            <% end if %>
            <input type="text" name="neotaxno" value="<%= ooffjungsan.FOneItem.Fneotaxno %>" size="20" maxlength="32" <%= CHKIIF(ISNULL(ooffjungsan.FOneItem.Ftaxlinkidx),"","class='text_ro' READONLY") %>>(TAXNO)
      	    <br>
      	    <input type="text" name="eseroEvalSeq" value="<%= ooffjungsan.FOneItem.FeseroEvalSeq %>" size="30" maxlength="30" <%= CHKIIF(ISNULL(ooffjungsan.FOneItem.Ftaxlinkidx),"","class='text_ro' READONLY") %> >(이세로 승인번호)

        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6A6A6" >추가정보</td>
        <td>

        <input type="checkbox" name="ispreFixTaxDateForce" <%= CHKIIF(isNULL(ooffjungsan.FOneItem.FpreFixedTaxDate),"","checked") %> >발행일강제지정
        <input type="text" name="preFixedTaxDate" value="<%= ooffjungsan.FOneItem.FpreFixedTaxDate %>" size="10" maxlength="10">
        (연동발행일경우 지정시 사용)

        <% if not isNULL(ooffjungsan.FOneItem.FrefPayreqIdx) then %>
        <input type="hidden" name="refPayreqIdx" value="<%= ooffjungsan.FOneItem.FrefPayreqIdx %>" >
        <b>결제요청서 IDX : <%= ooffjungsan.FOneItem.FrefPayreqIdx %></b>
        <% end if %>
        <!--
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="checkbox" name="isrefPay" <%= CHKIIF(isNULL(ooffjungsan.FOneItem.FrefPayreqIdx),"","checked") %> >결제요청서로 결제
        &nbsp; 결제요청IDX
        <input type="text" name="refPayreqIdx" value="<%= ooffjungsan.FOneItem.FrefPayreqIdx %>" size="7" maxlength="9">
        -->
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >입금일</td>
        <td>
            <input type="text" name="ipkumdate" value="<%= ooffjungsan.FOneItem.Fipkumdate %>" size="10" maxlength="10">
            <a href="javascript:calendarOpen(frmMaster.ipkumdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >그룹코드</td>
        <td>
            <input type="text" class="text" name="groupid" value="<%= ooffjungsan.FOneItem.Fgroupid %>" size="10" >
      	<input type="button" class="button" value="Code검색" onclick="popSearchGroupID(this.form.name,'groupid');" >
      	<input type="button" value="저장" onclick="saveGroupid(frmMaster);" <%= chkIIF(ooffjungsan.FOneItem.Ffinishflag>1,"disabled","") %> >
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6A6A6" >계정과목코드</td>
        <td>
            
            <input type="text" class="text_ro" value="<%= ooffjungsan.FOneItem.Fjacc_nm %>" size="10" readonly >
            <input type="text" class="text" name="jacctcd" value="<%= ooffjungsan.FOneItem.Fjacctcd %>" size="7" >
      	    <input type="button" value="저장" onclick="saveJacct(frmMaster);" >
      	    <!-- 기본 계정과목(미 입력시)은 [매입-상품매출원가,매출-] -->
        </td>
    </tr>
    <!--
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >네오포트발행</td>
        <td>
           <input type="checkbox" name="availneoport" <%= CHKIIF(ooffjungsan.FOneItem.Favailneo=1,"checked","") %>>가능
      	   <input type="button" value="저장" onclick="saveAvailNeo(frmMaster);" <%= chkIIF(ooffjungsan.FOneItem.Ffinishflag>=3,"disabled","") %> >
        </td>
    </tr>
    -->
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >총상품수</td>
        <td>
            <%= ooffjungsan.FOneItem.Ftot_itemno %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >총판매액</td>
        <td>
            <%= FormatNumber(ooffjungsan.FOneItem.Ftot_realsellprice,0) %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >실정산액</td>
        <td>

            <%= FormatNumber(ooffjungsan.FOneItem.Ftot_jungsanprice,0) %>
            <% if ooffjungsan.FOneItem.Ftot_realsellprice<>0 then %>
        		(<%= CLng((ooffjungsan.FOneItem.Ftot_realsellprice-ooffjungsan.FOneItem.Ftot_jungsanprice)/ooffjungsan.FOneItem.Ftot_realsellprice*100*100)/100 %> %)
        	<% end if %>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#E6E6E6" >비고</td>
        <td>
            <textarea name="comment" cols="70" rows="6"><%= ooffjungsan.FOneItem.Fcomment %></textarea>
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td colspan="2" align="center">
        <input type="button" value=" 기타 정보 변경 " onclick="SaveFrm(frmMaster,'masteretcedit');">
        </td>
    </tr>
    </form>
</table>

<%
set opartner =   Nothing
set ooffjungsan = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->