<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kaffa/itemsalecls.asp"-->
<%
Dim discountKey					'할인 인덱스키
Dim discountTitle				'할인제목
Dim promotionType				'프로모션타입..여기선 0으로 하기
Dim stDT						'시작일
Dim edDT						'종료일
Dim discountPro					'할인율
Dim discountbuyRule				'매입가구분(0:매입가지정, 1:판매가의N%)
Dim discountbuyPro				'판매가의N%
Dim regdate						'등록일
Dim lastupdate					'최근수정일
Dim openDate					'??
Dim expiredDate					'??
Dim regUserID					'등록자ID
Dim lastUpUserID				'최근수정자ID
Dim discountStatus, discountStatusStr
discountKey 	= request("discountKey")
discountStatus	= 0
Dim clsSale, sMode
sMode  = "I"
If discountKey <> "" Then
	Set clsSale = new CSale
		sMode  = "U"
		clsSale.FRectDiscountKey = discountKey
		clsSale.fnGetSaleConts

		discountTitle		= clsSale.FOneItem.FDiscountTitle
		promotionType		= clsSale.FOneItem.FPromotionType
		stDT				= clsSale.FOneItem.FStDT
		edDT				= clsSale.FOneItem.FEdDT
		discountPro			= clsSale.FOneItem.FDiscountPro
		discountbuyRule		= clsSale.FOneItem.FDiscountbuyRule
		discountbuyPro		= clsSale.FOneItem.FDiscountbuyPro
		regdate				= clsSale.FOneItem.FRegdate
		lastupdate			= clsSale.FOneItem.FLastupdate
		openDate			= clsSale.FOneItem.FOpenDate
		expiredDate			= clsSale.FOneItem.FExpiredDate
		regUserID			= clsSale.FOneItem.FRegUserID
		lastUpUserID		= clsSale.FOneItem.FLastUpUserID
		discountStatus		= clsSale.FOneItem.getDiscountStatus
		discountStatusStr   = clsSale.FOneItem.getSaleStateStr
	Set clsSale = nothing
End If
%>
<script language="javascript">
function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function jsChSetValue(iVal){
	if(iVal ==1){
		document.all.divM.style.display = "";
	}else{
		document.all.divM.style.display = "none";
	}
}
function jsSubmitSale(){
	var frm = document.frmReg;
    <% if (discountStatus=9) then %>
        alert('종료된 내역은 수정 불가 합니다.');
        return;
    <% end if %>

	if(!frm.discountTitle.value){
		alert("제목을 입력해 주세요");
		frm.discountTitle.focus();
		return ;
	}

	if(!frm.stDT.value ){
	  	alert("시작일을 입력해주세요");
	  	frm.stDT.focus();
	  	return ;
  	}

  	if(!frm.edDT.value ){
	  	alert("종료일을 입력해주세요");
	  	frm.edDT.focus();
	  	return ;
  	}

  	if(frm.edDT.value){
	  	if(frm.stDT.value > frm.edDT.value){
		  	alert("종료일이 시작일보다 빠릅니다. 다시 입력해주세요");
		  	frm.edDT.focus();
		  	return ;
	  	}
	}

	var nowDate = "<%=date()%>";
	if (frm.discountStatus.value!=9){
    	if(frm.stDT.value < nowDate){
    		alert("시작일이 오픈(핸재)일보다  빠르면 안됩니다. 시작일을 다시 선택해주세요");
    		frm.stDT.focus();
    		return ;
    	}
    }
  	if(!frm.discountPro.value){
		alert("할인율을 입력해 주세요");
		frm.discountPro.focus();
		return ;
	}
	if(confirm("저장 하시겠습니까?")){
		frm.submit();
	}
}
</script>
<table width="900" border="0" align="left" class="a" cellpadding="3" cellspacing="1"  >
<form name="frmReg" method="post" action="saleProc.asp" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="discountKey" value="<%=discountKey%>">
<input type="hidden" name="sMode" value="<%=sMode%>">
<tr>
	<td>
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">제목</td>
			<td bgcolor="#FFFFFF"><input type="text" name="discountTitle" size="30" maxlength="64" value="<%=discountTitle%>"></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">기간</td>
			<td bgcolor="#FFFFFF">
				시작일 : <input type="text" name="stDT" size="10" onClick="jsPopCal('stDT');"  style="cursor:hand;" value="<%=chkiif(stDT<>"",LEFT(stDT,10),date())%>" >
				~ 종료일 : <input type="text" name="edDT" size="10" onClick="jsPopCal('edDT');" style="cursor:hand;" value="<%=LEFT(edDT,10)%>">
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">할인율</td>
			<td bgcolor="#FFFFFF"><input type="text" name="discountPro" size="4" style="text-align:right;" value="<%=discountPro%>">%</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">매입가구분</td>
			<td bgcolor="#FFFFFF">
				<select name= "discountbuyRule" onchange="jsChSetValue(this.value);" class="select">
					<option value="0" <%=chkiif(discountbuyRule="0","selected","")%> >매입가지정</option>
					<option value="1" <%=chkiif(discountbuyRule="1","selected","")%>>판매가의N%</option>
				</select>
				<span id="divM" style="display:<%IF discountbuyRule<> 1 THEN %>none<%END IF%>;">&nbsp;&nbsp;판매가의<input type="text" size="4" name="discountbuyPro" maxlength="10" value="<%=discountbuyPro%>" style="text-align:right;">%</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">상태</td>
			<td bgcolor="#FFFFFF">
			    <%= discountStatusStr %>
				<% If discountStatus = 0 Then %>
				   <input type="checkbox" name="discountStatus" value="7">오픈요청
				<% ElseIf discountStatus = 6 or discountStatus = 7 Then %>
					<input type="checkbox" name="discountStatus" value="9">종료요청
			    <% else %>
			        <input type="hidden" name="discountStatus" value="<%=discountStatus%>">
				<% End If %>

				<% If Not isNULL(expiredDate) Then %>
				(종료일 : <%=expiredDate%>)
				<% End If %>
			</td>
			<td bgcolor="#FFFFFF" colspan="2"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">
		<a href="javascript:jsSubmitSale();"><img src="/images/icon_save.gif"  border="0"></a>
		<a href="javascript:history.back();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->