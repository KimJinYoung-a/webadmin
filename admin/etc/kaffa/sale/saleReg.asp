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
Dim discountKey					'���� �ε���Ű
Dim discountTitle				'��������
Dim promotionType				'���θ��Ÿ��..���⼱ 0���� �ϱ�
Dim stDT						'������
Dim edDT						'������
Dim discountPro					'������
Dim discountbuyRule				'���԰�����(0:���԰�����, 1:�ǸŰ���N%)
Dim discountbuyPro				'�ǸŰ���N%
Dim regdate						'�����
Dim lastupdate					'�ֱټ�����
Dim openDate					'??
Dim expiredDate					'??
Dim regUserID					'�����ID
Dim lastUpUserID				'�ֱټ�����ID
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
        alert('����� ������ ���� �Ұ� �մϴ�.');
        return;
    <% end if %>

	if(!frm.discountTitle.value){
		alert("������ �Է��� �ּ���");
		frm.discountTitle.focus();
		return ;
	}

	if(!frm.stDT.value ){
	  	alert("�������� �Է����ּ���");
	  	frm.stDT.focus();
	  	return ;
  	}

  	if(!frm.edDT.value ){
	  	alert("�������� �Է����ּ���");
	  	frm.edDT.focus();
	  	return ;
  	}

  	if(frm.edDT.value){
	  	if(frm.stDT.value > frm.edDT.value){
		  	alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		  	frm.edDT.focus();
		  	return ;
	  	}
	}

	var nowDate = "<%=date()%>";
	if (frm.discountStatus.value!=9){
    	if(frm.stDT.value < nowDate){
    		alert("�������� ����(����)�Ϻ���  ������ �ȵ˴ϴ�. �������� �ٽ� �������ּ���");
    		frm.stDT.focus();
    		return ;
    	}
    }
  	if(!frm.discountPro.value){
		alert("�������� �Է��� �ּ���");
		frm.discountPro.focus();
		return ;
	}
	if(confirm("���� �Ͻðڽ��ϱ�?")){
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
			<td width="100" bgcolor="<%= adminColor("tabletop") %>" align="center">����</td>
			<td bgcolor="#FFFFFF"><input type="text" name="discountTitle" size="30" maxlength="64" value="<%=discountTitle%>"></td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">�Ⱓ</td>
			<td bgcolor="#FFFFFF">
				������ : <input type="text" name="stDT" size="10" onClick="jsPopCal('stDT');"  style="cursor:hand;" value="<%=chkiif(stDT<>"",LEFT(stDT,10),date())%>" >
				~ ������ : <input type="text" name="edDT" size="10" onClick="jsPopCal('edDT');" style="cursor:hand;" value="<%=LEFT(edDT,10)%>">
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">������</td>
			<td bgcolor="#FFFFFF"><input type="text" name="discountPro" size="4" style="text-align:right;" value="<%=discountPro%>">%</td>
			<td width="100"  bgcolor="<%= adminColor("tabletop") %>"  align="center">���԰�����</td>
			<td bgcolor="#FFFFFF">
				<select name= "discountbuyRule" onchange="jsChSetValue(this.value);" class="select">
					<option value="0" <%=chkiif(discountbuyRule="0","selected","")%> >���԰�����</option>
					<option value="1" <%=chkiif(discountbuyRule="1","selected","")%>>�ǸŰ���N%</option>
				</select>
				<span id="divM" style="display:<%IF discountbuyRule<> 1 THEN %>none<%END IF%>;">&nbsp;&nbsp;�ǸŰ���<input type="text" size="4" name="discountbuyPro" maxlength="10" value="<%=discountbuyPro%>" style="text-align:right;">%</span>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>"  align="center">����</td>
			<td bgcolor="#FFFFFF">
			    <%= discountStatusStr %>
				<% If discountStatus = 0 Then %>
				   <input type="checkbox" name="discountStatus" value="7">���¿�û
				<% ElseIf discountStatus = 6 or discountStatus = 7 Then %>
					<input type="checkbox" name="discountStatus" value="9">�����û
			    <% else %>
			        <input type="hidden" name="discountStatus" value="<%=discountStatus%>">
				<% End If %>

				<% If Not isNULL(expiredDate) Then %>
				(������ : <%=expiredDate%>)
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