<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ csv �߰�
' Hieditor : 2017.06.27 ������ ����
'			 2017.06.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim page, mode, designer,suplyer,shopid,dataarr,bufarr,bufstr,odataarr, buforderno, pbrandid, loginsite, idx
dim i, shopsuplycash, buycash, j
dim ttlitemno : ttlitemno=0
	shopid   = requestCheckvar(request("shopid"),32)
	page     = requestCheckvar(request("page"),10)
	mode     = requestCheckvar(request("mode"),32)
	designer = requestCheckvar(request("designer"),32)
	suplyer  = requestCheckvar(request("suplyer"),32)
	dataarr  = request("dataarr")
	odataarr = request("dataarr")
	pbrandid = requestCheckvar(request("pbrandid"),32)
	idx      = requestCheckvar(request("idx"),10)

if suplyer<>"10x10" then designer = suplyer
if page="" then page=1
if mode="" then mode="bybrand"

dim ioffitem
set ioffitem  = new COffShopItem
	ioffitem.FPageSize = 100
	ioffitem.FCurrPage = page
	ioffitem.FRectDesigner = designer

	if suplyer="10x10" then
		ioffitem.FRectShopid = shopid
		ioffitem.FRectDesignerjungsangubun = "'2','4','5'"
	else
		ioffitem.FRectShopid = shopid
		ioffitem.FRectDesignerjungsangubun = "'6','8'"
	end if

	if (dataarr<>"") then
		dataarr = split(dataarr,vbcrlf)   ''�� ���ͺи�
		for i=LBound(dataarr) to UBound(dataarr)
		    
		    if (Trim(dataarr(i))<>"") then
			bufarr = split(dataarr(i),chr(9))  ''�÷� �Ǻи� 
	    		if UBound(bufarr)=1 then       '' barcode(/t)����
	    		    bufstr = bufstr + "'" + requestCheckvar(bufarr(0) ,20) + "'" + ","
	    	    end if
		    end if
		next
	    
	    bufstr = Trim(bufstr)
	    if (Right(bufstr,1)=",") then
		    bufstr = Left(bufstr,Len(bufstr)-1)
		end if


		ioffitem.FRectDesigner = pbrandid
		ioffitem.FRectBarCodeArr = bufstr
		
		''response.write bufstr
		if (bufstr<>"") then
		    ioffitem.GetOffLineJumunByArr

		    loginsite = ioffitem.Floginsite
	    end if
	end if
%>
<script type="text/javascript">

function enablebrand(bool){
	//document.frm.designer.disabled = bool;
}

function search(frm){
	frm.submit();
}

function AddArr(){
	var upfrm = document.frmArrupdate;
	var frm;
	var unreg="";
	var pass = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

	if (!pass) {
		alert('���� �������� �����ϴ�.');
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.itemnamearr.value = "";
	upfrm.itemoptionnamearr.value = "";
	upfrm.designerarr.value = "";
	upfrm.foreign_sellcasharr.value = "";
	upfrm.foreign_suplycasharr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){

				if (!IsInteger(frm.itemno.value)){
					alert('������ ������ �����մϴ�.');
					frm.itemno.focus();
					return;
				}

				if (frm.itemno.value=="0"){
					alert('������ �Է��ϼ���.');
					frm.itemno.focus();
					return;
				}

				if(frm.foreign_sellcash.value==0&&(document.frm.loginsite.value=="WSLWEB")){
					if ( unreg == "" ){
						unreg = frm.itemid.value;
					}else{
				 		unreg = unreg + "," + frm.itemid.value;
					}

					//�̵�� ��ǰ�� �ϴ� �ִ´�. ���Ŀ� �ؿܻ�ǰ�ܿ� ��ǰ������ ����� �Է¾ȵǰ� ������ �ؾ���.	//2017.06.12 �ѿ��
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
					upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
					upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
					upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value+frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value+frm.foreign_suplycash.value + "|";
					upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
					upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
					upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				}else{
					upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
					upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
					upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
					upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
					upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
					upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
					upfrm.foreign_sellcasharr.value = upfrm.foreign_sellcasharr.value+frm.foreign_sellcash.value + "|";
					upfrm.foreign_suplycasharr.value = upfrm.foreign_suplycasharr.value+frm.foreign_suplycash.value + "|";
					upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
					upfrm.itemnamearr.value = upfrm.itemnamearr.value + frm.itemname.value + "|";
					upfrm.itemoptionnamearr.value = upfrm.itemoptionnamearr.value + frm.itemoptionname.value + "|";
					upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
				}
			}
		}
	}

	if (unreg!=""){
		alert("�����Ͻ� ��ǰ �� ��ǰ�ڵ� ["+unreg+"]�� �̵�ϻ�ǰ�Դϴ�. \n��ǰ ��� �� �ֹ����ּ���");
	}

	opener.ReActItems('<%= idx %>',upfrm.itemgubunarr.value,upfrm.itemarr.value,upfrm.itemoptionarr.value,
		upfrm.sellcasharr.value,upfrm.suplycasharr.value,upfrm.buycasharr.value,upfrm.itemnoarr.value,upfrm.itemnamearr.value,
		upfrm.itemoptionnamearr.value,upfrm.designerarr.value,upfrm.foreign_sellcasharr.value,upfrm.foreign_suplycasharr.value);

}
</script>
<table width="840" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="frm" method="post" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="suplyer" value="<%= suplyer %>" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="loginsite" value="<%=loginsite%>">
<tr>
	<td class="a" >
		�귣��ID : <%	drawSelectBoxDesignerWithName "pbrandid", pbrandid %> <!-- input type="text" name="pbrandid" value="<%= pbrandid %>" -->
		
		<br>
		<textarea name="dataarr" cols=80 rows=6><%= odataarr %></textarea> <input type= button value=clear onclick="frm.dataarr.value=''; frm.pbrandid.value=''">
	</td>
	<td class="a" align="right">
		<a href="javascript:search(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
<tr>
    <td class="a">�ִ� 200�� <br> ������ �и� ���ڵ�(��)����(����)  <br>������, �����Ź�� �˻�����</td>
</tr>
</form>
</table>

<table width="840" cellspacing="1" class="a" bgcolor=#3d3d3d>
<% if ioffitem.FresultCount>0 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="right">�ѰǼ�: <%= ioffitem.FResultCount %> &nbsp; </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="right"><input type="button" value="���� ������ �߰�" onclick="AddArr()"></td>
</tr>
<% end if %>
<tr bgcolor="#DDDDFF">
	<td width="20"><input type="checkbox" name="ckall" onClick="AnSelectAllFrame(this)"></td>
	<td width="50">�̹���</td>
	<td width="50">�귣��ID</td>
	<td width="80">BarCode</td>
	<td width="100">��ǰ��</td>
	<td width="80">�ɼǸ�</td>
	<td width="60">�ǸŰ�</td>
	<td width="60">���ް�</td>
	<td width="48">���޸���</td>
	<td width="50">����</td>
	<td width="70">���</td>
</tr>
<% for i=0 to ioffitem.FResultCount -1 %>
<%
if session("ssBctDiv")="502" or session("ssBctDiv")="503" then
	shopsuplycash = ioffitem.FItemList(i).GetFranchiseSuplycash
	buycash		  = ioffitem.FItemList(i).GetFranchiseBuycash
else
	shopsuplycash = ioffitem.FItemList(i).GetOfflineSuplycash
	buycash		  = ioffitem.FItemList(i).GetFranchiseBuycash
end if
%>
<form name="frmBuyPrc_<%= i %>" >
<input type="hidden" name="itemgubun" value="<%= ioffitem.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ioffitem.FItemList(i).Fshopitemid %>">
<input type="hidden" name="itemoption" value="<%= ioffitem.FItemList(i).Fitemoption %>">
<input type="hidden" name="itemname" value="<%= ioffitem.FItemList(i).FShopItemName %>">
<input type="hidden" name="itemoptionname" value="<%= ioffitem.FItemList(i).FShopItemOptionName %>">
<input type="hidden" name="desingerid" value="<%= ioffitem.FItemList(i).FMakerid %>">
<input type="hidden" name="sellcash" value="<%= ioffitem.FItemList(i).Fshopitemprice %>">
<input type="hidden" name="suplycash" value="<%= shopsuplycash %>">
<input type="hidden" name="buycash" value="<%= buycash %>">
<input type="hidden" name="foreign_sellcash" value="<%= ioffitem.FItemList(i).Fforeign_sellprice %>">
<input type="hidden" name="foreign_suplycash" value="<%= ioffitem.FItemList(i).Fforeign_suplyprice %>">
<tr bgcolor="<%=CHKIIF(ioffitem.FItemList(i).Fisusing="N","#CCCCCC","#FFFFFF")%>">
	<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
	<td ><img src="<%= ioffitem.FItemList(i).FimageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
	<td ><%= ioffitem.FItemList(i).FMakerid %></td>
	<td ><%= ioffitem.FItemList(i).GetBarCode %></td>
	<td ><%= ioffitem.FItemList(i).FShopItemName %></td>
	<td ><%= ioffitem.FItemList(i).FShopItemOptionName %></td>
	<td align=right>
		<%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %>
	
		<br><font color="Gray"><%= ioffitem.fcurrencyChar %>&nbsp;<%= FormatNumber(ioffitem.FItemList(i).Fforeign_sellprice,2) %></font>
	</td>
	<td align=right>
		<%= FormatNumber(shopsuplycash,0) %>
	
		<br><font color="gray"><%= ioffitem.fcurrencyChar %>&nbsp;<%= FormatNumber(ioffitem.FItemList(i).Fforeign_suplyprice,2) %></font>
	</td>
	<td align=center>
	<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
	<%= 100-(CLng(shopsuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %
	<% end if %>
	</td>
	<%
	buforderno = 0
	for j=LBound(dataarr) to UBound(dataarr)
		bufarr = split(dataarr(j),chr(9))
		if UBound(bufarr)=1 then
			if (TRIM(bufarr(0))=TRIM(ioffitem.FItemList(i).Ftnbarcode)) or (TRIM(bufarr(0))=TRIM(ioffitem.FItemList(i).Fextbarcode)) then
				buforderno = bufarr(1)
			end if
		end if
	next
	
	ttlitemno=ttlitemno+buforderno
	%>
	<td ><input type="text" name="itemno" value="<%= buforderno %>" size="4" maxlength="4"></td>

	<td >
	    
	<% if ioffitem.FItemList(i).Fisusing="N" then %>
	<font color="red">�������</font><br>
	<% end if %>
	<% if ioffitem.FItemList(i).Foptusing="N" then %>
	<font color="red">ON�ɼ�X</font><br>
	<% end if %>
	<% if ioffitem.FItemList(i).IsSoldOut then %>
	<font color="red">ON�Ǹ�����</font><br>
	<% end if %>
	<% if ioffitem.FItemList(i).Flimityn="Y" then %>
	<font color="blue">ON����(<%= ioffitem.FItemList(i).getLimitNo %>)</font>
	<% end if %>
	</td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="11" align="right">
	 �Ѽ��� : <%=ttlitemno%>
	</td>
</tr>
</table>
<form name="frmArrupdate" method="post" action="">
	<input type="hidden" name="mode" value="arrins">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="foreign_sellcasharr" value="">
	<input type="hidden" name="foreign_suplycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="itemnamearr" value="">
	<input type="hidden" name="itemoptionnamearr" value="">
	<input type="hidden" name="designerarr" value="">
</form>

<%
set ioffitem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->