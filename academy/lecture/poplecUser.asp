<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' History : 2010.10.19 eastone 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/lecturer/lecUserCls.asp"-->
<%
Dim lecturer_id : lecturer_id = requestCheckVar(request("lecturer_id"),32)
Dim tenID : tenID = requestCheckVar(request("tenID"),32)

if (tenID="") then 
    tenID=lecturer_id
end if

dim olecuter 
set olecuter = new CLecUser
olecuter.FRectLecturerID = lecturer_id
olecuter.getOneLecUserInfo

dim otenInfo
set otenInfo = new CLecUser
otenInfo.FRectLecturerID = tenID
otenInfo.getTenLecUserInfo

%>
<script language='javascript'>
function chkComp(comp){
    var frm = comp.form;
    if (comp.value=="9"){
        frm.DefaultFreebeasongLimit.style.background = '#FFFFFF';
        frm.DefaultDeliveryPay.style.background  = '#FFFFFF';
        
        frm.DefaultFreebeasongLimit.readOnly = false;
        frm.DefaultDeliveryPay.readOnly = false;
        
        frm.DefaultFreebeasongLimit.value=frm.pDFL.value;
        frm.DefaultDeliveryPay.value=frm.pDDP.value;
        
        
    }else{
        frm.DefaultFreebeasongLimit.style.background = '#BBBBBB';
        frm.DefaultDeliveryPay.style.background  = '#BBBBBB';
        
        frm.DefaultFreebeasongLimit.readOnly = true;
        frm.DefaultDeliveryPay.readOnly = true;
        
        frm.DefaultFreebeasongLimit.value=0;
        frm.DefaultDeliveryPay.value=0;
    }
}

function clickDiy(comp){
    if (comp.value=="Y"){
        iDiyDlv.style.display="inline";
    }else{
        iDiyDlv.style.display="none";
    }
}

function saveInfo(frm){
    
    if (frm.lecturer_name.value.length<1){
        alert('강사(브랜드)명 을 입력하세요.');
        frm.lecturer_name.focus();
        return;
    }
    
    if (frm.en_name.value.length<1){
        alert('영문 표시명 을 입력하세요.');
        frm.en_name.focus();
        return;
    }
    
    if ((!frm.lec_yn[0].checked)&&(!frm.lec_yn[1].checked)){
        alert('강좌 진행 여부를 선택하세요.');
        frm.lec_yn[0].focus();
        return;
    }
    
    if ((!frm.diy_yn[0].checked)&&(!frm.diy_yn[1].checked)){
        alert('diy 진행 여부를 선택하세요.');
        frm.diy_yn[0].focus();
        return;
    }
    
    if (!IsDouble(frm.lec_margin.value)){
        alert('강좌 기본 마진은 숫자만 가능합니다.1~99');
        frm.lec_margin.focus();
        return;
    }
    
    if (!IsDouble(frm.diy_margin.value)){
        alert('DIY 기본 마진은 숫자만 가능합니다.1~99');
        frm.diy_margin.focus();
        return;
    }
    
    if ((frm.lec_yn[0].checked)&&((frm.lec_margin.value*1<1)||(frm.lec_margin.value*1>99))){
        alert('강좌 기본 마진을 입력하세요.1~99');
        frm.lec_margin.focus();
        return;
    }
    
    if ((frm.diy_yn[0].checked)&&((frm.diy_margin.value*1<1)||(frm.diy_margin.value*1>99))){
        alert('DIY 기본 마진을 입력하세요.1~99');
        frm.diy_margin.focus();
        return;
    }
    
    if ((frm.diy_yn[0].checked)&&(frm.diy_dlv_gubun.value.length<1)){
        alert('DIY 배송구분을 선택하세요.');
        frm.diy_dlv_gubun.focus();
        return;
    }
    
    if (frm.diy_dlv_gubun.value=="9"){
        if (!IsDigit(frm.DefaultFreebeasongLimit.value)){
            alert('배송비 기준 숫자만 가능합니다.');
            frm.DefaultFreebeasongLimit.focus();
            return;
        }
        
        if (!IsDigit(frm.DefaultDeliveryPay.value)){
            alert('배송비  숫자만 가능합니다.');
            frm.DefaultDeliveryPay.focus();
            return;
        }
        
        if (frm.DefaultFreebeasongLimit.value*1<=0){
            alert('금액을 0원 이상 입력하세요.');
            frm.DefaultFreebeasongLimit.focus();
            return;
        }
        
        if (frm.DefaultDeliveryPay.value*1<=0){
            alert('금액을 0원 이상 입력하세요.');
            frm.DefaultDeliveryPay.focus();
            return;
        }
        
    }
    
    if (confirm('저장하시겠습니까?')){
        frm.submit();
    }
}

function fnresearch(frm1,frm2){
    frm2.tenID.value = frm1.tenID.value;
    frm2.submit();
}

function getOnLoad(){
    chkComp(frmLecturer.diy_dlv_gubun);
    if (frmLecturer.diy_yn[0].checked){
        clickDiy(frmLecturer.diy_yn[0]);
    }
}

window.onload=getOnLoad;
</script>
<!--
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	강사 ID : <input type="text" name="lecturer_id" value="<%= lecturer_id %>" Maxlength="32" size="16">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
-->
<table width="700" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="reloadFrm" method="get">
<input type="hidden" name="lecturer_id" value="<%= lecturer_id %>">
<input type="hidden" name="tenID" value="<%= tenID %>">
</form>
<form name="frmLecturer" method="post" action="doLecUserEdit.asp">
	<tr >
		<td width="130" bgcolor="#DDDDFF">강사(브랜드)ID</td>
		
		<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="lecturer_id" value="<%= lecturer_id %>" size="28" maxlength="32" readonly class="text_ro">
		</td>
		<td bgcolor="#FFFFFF" width="180" align="center">
		<input type="text" name="tenID" value="<%= tenID %>" size="12"> 
		<input type="button" value="검색" onclick="fnresearch(frmLecturer,reloadFrm);">
		
		</td>
	</tr>
	
	<tr >
		<td width="120" bgcolor="#DDDDFF">강사(브랜드)명</td>
		
		<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="lecturer_name" value="<%= olecuter.FOneItem.Flecturer_name %>" size="28" maxlength="16" >
		</td>
		<td bgcolor="#FFFFFF" ><%= otenInfo.FOneItem.FTen_socname_kor %></td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF">영문 표시명</td>
		
		<td bgcolor="#FFFFFF" colspan="3">
		<input type="text" name="en_name" value="<%= olecuter.FOneItem.Fen_name %>" size="28" maxlength="16" >
		</td>
		<td bgcolor="#FFFFFF" ><%= otenInfo.FOneItem.FTen_socname %></td>
	</tr>
	
	<tr >
		<td width="120" bgcolor="#DDDDFF" rowspan="2">강좌 진행 여부</td>
		<td bgcolor="#FFFFFF" rowspan="2">
		<input type="radio" name="lec_yn" value="Y" <%= CHKIIF(olecuter.FOneItem.Flec_yn="Y","checked","") %> > Y
		<input type="radio" name="lec_yn" value="N" <%= CHKIIF(olecuter.FOneItem.Flec_yn="N","checked","") %> > N
		</td>
		<td width="120" bgcolor="#DDDDFF">강좌기본마진</td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="lec_margin" value="<%= olecuter.FOneItem.Flec_margin %>" size="4" maxlength="3"> (%)
		</td>
		<td bgcolor="#FFFFFF" ><%= otenInfo.FOneItem.FTen_defaultmargine %></td>
	</tr>
	<tr>
	    <td width="120" bgcolor="#DDDDFF">재료기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="mat_margin" value="<%= olecuter.FOneItem.Fmat_margin %>" size="4" maxlength="3"> (%)
		</td>
		<td bgcolor="#FFFFFF" ></td>
	</tr>
	<tr >
		<td width="120" bgcolor="#DDDDFF" >DIY 진행 여부</td>
		<td  bgcolor="#FFFFFF" width="200" >
		<input type="radio" name="diy_yn" value="Y" <%= CHKIIF(olecuter.FOneItem.Fdiy_yn="Y","checked","") %> onClick="clickDiy(this);"> Y
		<input type="radio" name="diy_yn" value="N" <%= CHKIIF(olecuter.FOneItem.Fdiy_yn="N","checked","") %> onClick="clickDiy(this);"> N
		</td>
		<td width="120" bgcolor="#DDDDFF">기본마진</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="diy_margin" value="<%= olecuter.FOneItem.Fdiy_margin %>" size="4" maxlength="3"> (%)
		</td>
		<td bgcolor="#FFFFFF" ></td>
	</tr>
	
	<tr id="iDiyDlv" style="display:none">
		<td width="120" bgcolor="#DDDDFF">DIY배송구분</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<select name="diy_dlv_gubun" onChange="chkComp(this);">
		<option value="0" <%= CHKIIF(olecuter.FOneItem.Fdiy_dlv_gubun<>9,"selected","") %>>기본(업체무료배송)
		<option value="9" <%= CHKIIF(olecuter.FOneItem.Fdiy_dlv_gubun=9,"selected","") %>>업체 조건배송
		</select>
		<br>
		<input type="hidden" name="pDFL" value="<%= olecuter.FOneItem.FDefaultFreebeasongLimit %>">
		<input type="hidden" name="pDDP" value="<%= olecuter.FOneItem.FDefaultDeliveryPay %>">
		<input type="text" name="DefaultFreebeasongLimit" value="<%= olecuter.FOneItem.FDefaultFreebeasongLimit %>" size="9" maxlength="9">원 이상 무료배송
		/미만 배송비 <input type="text" name="DefaultDeliveryPay" value="<%= olecuter.FOneItem.FDefaultDeliveryPay %>" size="9" maxlength="9">원
		</td>
		<td bgcolor="#FFFFFF" >
		<%= otenInfo.FOneItem.getTenDlvStr %>
		</td>
	</tr>
	
	<tr >
		<td width="120" bgcolor="#DDDDFF">등록일</td>
		<td colspan="3" bgcolor="#FFFFFF">
		<%= olecuter.FOneItem.Fregdate %>
		</td>
		<td bgcolor="#FFFFFF" >
		<%= olecuter.FOneItem.getTenDlvStr %>
		</td>
	</tr>
	<tr>
		<td colspan="5" bgcolor="#FFFFFF" height="25" align="center">
		<input type="button" value="저 장" onClick="saveInfo(frmLecturer)">
		</td>
	</tr>
</form>
</table>
<%
set olecuter = Nothing
set otenInfo = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
