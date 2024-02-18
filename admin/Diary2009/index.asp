<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���̾ ����Ʈ ����
' History : 2015.09.14 ���¿� ����(���ù�ȣ,��뿩�� �ϰ�ó�� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/classes/DiaryCls.asp"-->
<%
dim limited, itemid
dim CateCode , yearUse , isusing, mdpick

dim sBrand,arrItemid
CateCode = request("cate")
yearUse = "2009"
isusing = request("isusingbox")
sBrand = request("ebrand")
arrItemid = request("aitem")
mdpick = request("mdpick")
limited = request("limited")
itemid      = requestCheckvar(request("itemid"),255)
dim page , i
	page = requestCheckVar(request("page"),5)
	if page = "" then page = 1

if itemid<>"" then
	dim iA ,arrTemp',arrItemid
  	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oDiary
set oDiary = new DiaryCls
	oDiary.FPageSize = 50
	oDiary.FCurrPage = page
	oDiary.frectcate = CateCode
	oDiary.frectisusing = isusing
	oDiary.FrectMakerid = sBrand
	oDiary.FRectArrItemid = arrItemid
	oDiary.frectmdpick = mdpick
	oDiary.frectlimited = limited
	oDiary.getDiaryList

%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">

//�ű� ��� �˾�
function popRegNew(){
	var popRegNew = window.open('/admin/diary2009/DiaryReg.asp','popRegNew','width=600,height=750,status=yes')
	popRegNew.focus();
}

function popRegItems(){
	var popRegItems = window.open('/admin/diary2009/lib/pop_itemAddinfo.asp','popRegItems','width=1000,height=750,scrollbars=yes,resizable=yes')
	popRegItems.focus();
}

//���� �˾�
function popRegModi(idx){
	var popRegModi = window.open('/admin/diary2009/DiaryReg.asp?mode=edit&id='+ idx,'popRegModi','width=600,height=750')
	popRegModi.focus();
}

//�̹��� ����
function contents_option(){
	var contents_option = window.open('/admin/diary2009/imagemake/imagemake_list.asp','contents_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	contents_option.focus();
}

//����� �귣�� ����(20160907���¿�)
function fnspbrand(){
	var spbrand = window.open('/admin/diary2009/specialbrand/specialbrand_list.asp','spbrand','width=1024,height=768,scrollbars=yes,resizable=yes');
	spbrand.focus();
}

function keyword_option(){
	var keyword_option = window.open('/admin/diary2009/option/keyword_option.asp','keyword_option','width=1024,height=768,scrollbars=yes,resizable=yes');
	keyword_option.focus();
}

function detail_view(DiaryID){
	var detail_view = window.open('/admin/diary2009/option/detail_option.asp?DiaryID='+DiaryID,'detail_view','width=1024,height=768,scrollbars=yes,resizable=yes');
	detail_view.focus();
}

function edit(id){
	document.location.href="/admin/diary2009/DiaryReg.asp?mode=edit&id="+id;
}

//���� ���� ������ �߰�,���� �˾�
function popInfoReg(idx){
	var popInfoReg = window.open('/admin/diary2009/option/pop_diary_info_reg.asp?mode=modify&diaryid=' + idx,'popInfoReg','width=620,height=800,status=no,resizable=yes,scrollbars=yes')
	popInfoReg.focus();
}

//�� ���� ������ �߰�,���� �˾�
function popContReg(idx){
	alert('������');
	var popContReg = window.open('/admin/diary2009/pop_diary_cont_reg.asp?mode=modify&diaryid=' + idx,'popContReg','width=620,height=800,resizable=yes,scrollbars=yes')
	popContReg.focus();
}

//���ΰ���
function popMainReg(){
	var popMainReg = window.open('/admin/diary2009/pop_DiaryMain_reg.asp','popContReg','width=620,height=800,resizable=yes,scrollbars=yes')
	popMainReg.focus();
}

//�̺�Ʈ����
function popeventReg(){
	var popeventReg = window.open('/admin/diary2009/event.asp','popeventReg','width=1024,height=768,resizable=yes,scrollbars=yes')
	popeventReg.focus();
}

//1+1����
function pop1plus1Reg(){
	var pop1plus1Reg = window.open('/admin/diary2009/diary_OneplusOne.asp','pop1plus1Reg','width=1024,height=768,resizable=yes,scrollbars=yes')
	pop1plus1Reg.focus();
}

// mdpick ���� ���� �˾�
function popMdpickSort(){
	var popMdpickSort = window.open('/admin/diary2009/diary_mdpicksort.asp','popMdpickSort','width=1024,height=768,resizable=yes,scrollbars=yes')
	popMdpickSort.focus();
}

//�귣�����ͺ����
function popBrandInterview(){
	var popBrandInterview = window.open('/admin/diary2009/brand_interview.asp','popBrandInterview','width=1024,height=768,resizable=yes,scrollbars=yes')
	popBrandInterview.focus();
}

//���̾ �������̹��� ���/����
function popPrevImg(idx){
	var popPrevImgAction = window.open('/admin/diary2009/PreviewImg.asp?idx='+idx+'','popPrevImgAction','width=1024,height=768,resizable=yes,scrollbars=yes')
	popPrevImgAction.focus();
}

//MD's Pick ���������ϱ�
function popMainMDPickReg(){
	<% If oDiary.FTotalCount < 18 Then %>
	alert("�� 18�� �̻��� �־�� MD's Pick�� ������ �˴ϴ�.");
	return;
	<% End If %>

	var popMainMDPickReg;
	popMainMDPickReg = window.open("<%=wwwUrl%>/chtml/diary/main_make_xml.asp?imagecount=18", "popMainMDPickReg","width=800,height=600,scrollbars=yes,resizable=yes");
	popMainMDPickReg.focus();
}

//mdpick ��ü����
var ichk;
ichk = 1;
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.name=="chkI")){
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
			}
		}
	}
}

//mdpick �ϰ� ����
function jsSortIsusing() {
	if (confirm("MDpick��뿩�θ� ���� �Ͻðڽ��ϱ�?") == true){    //Ȯ��
		var frm;
		var sValue;
		var mdpick;
		frm = document.fitem;
		sValue = ""; //idx
		sCheck = ""; //mdpick o,x
		chkSel	= 0;
	
		if (frm.chkI.length > 1){
			for (var i=0;i<frm.chkI.length;i++){
				if(frm.chkI[i].checked) chkSel++;
	
				if (frm.chkI[i].checked){
					if (sValue==""){
						sValue = frm.chkI[i].value;
					}else{
						sValue =sValue+","+frm.chkI[i].value;
					}
					
					frm.mdpickchk[i].value="o";
					if (sCheck==""){
						sCheck = frm.mdpickchk[i].value;
					}else{
						sCheck =sCheck+","+frm.mdpickchk[i].value;
					}
				}else{
					if (sValue==""){
						sValue = frm.chkI[i].value;
					}else{
						sValue =sValue+","+frm.chkI[i].value;
					}
					frm.mdpickchk[i].value="x";
					if (sCheck==""){
						sCheck = frm.mdpickchk[i].value;
					}else{
						sCheck =sCheck+","+frm.mdpickchk[i].value;
					}
				}
			}
		}else{
			if(frm.chkI.checked) chkSel++;
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
				sCheck = frm.mdpickchk.value;
			}
		}
		document.frmreg.mdpick.value = sCheck;
		document.frmreg.did.value = sValue;
		document.frmreg.submit();
	}else{
	    return;
	}
}

//�������,��뿩�� ����
function jsjunsiSortIsusing() {

	var frm;
	var sValue, sortNo, isusing;
	frm = document.fitem;
	sValue = "";
	sortNo = "";
	isusing = "";
	chkSel	= 0;

	if (frm.chkJ.length > 1){
		for (var i=0;i<frm.chkJ.length;i++){
			if(frm.chkJ[i].checked) chkSel++;

			if (frm.isusing[i].value ==''){
				alert('��뿩�θ� �����ϼ���.');
				frm.isusing[i].focus();
				return;
			}
			if (frm.chkJ[i].checked){
				if (sValue==""){
					sValue = frm.chkJ[i].value;
				}else{
					sValue =sValue+","+frm.chkJ[i].value;
				}

				// �������
				if (sortNo==""){
					sortNo = frm.sortNo[i].value;
				}else{
					sortNo =sortNo+","+frm.sortNo[i].value;
				}

				// ��뿩��
				if (isusing==""){
					isusing = frm.isusing[i].value;
				}else{
					isusing =isusing+","+frm.isusing[i].value;
				}
			}
		}
	}else{
		if(frm.chkJ.checked) chkSel++;
		if(frm.chkJ.checked){
			sValue = frm.chkJ.value;
			if(!IsDigit(frm.sortNo.value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sortNo.focus();
				return;
			}
			sortNo 	=  frm.sortNo.value;
			isusing =  frm.isusing.value;
		}
	}
	if(chkSel<=0) {
		alert("������ ���̾ �����ϴ�.");
		return;
	}
	document.frmSortIsusing.sortnoarr.value = sortNo;
	document.frmSortIsusing.isusingarr.value = isusing;
	document.frmSortIsusing.detailidxarr.value = sValue;
	document.frmSortIsusing.submit();
}

//���ù�ȣ,��뿩�� ���ÿ��� ��ü����
var ichk;
ichk = 1;
function jsjunsiChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	if(!frm.chkJ) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		var e = frm.elements[i];
		if ((e.name=="chkJ")){
			if ((e.type=="checkbox")) {
				e.checked = blnChk ;
			}
		}
	}
}

//��뿩�� ��ü ����
function jsIsusingChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkJ.length > 1){
		for (var i=0;i<frm.isusing.length;i++){
			frm.isusing[i].value=selv;
		}
	}else{
		frm.isusing.value=selv;
	}
}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<!-- ��� �˻��� ���� -->
		<div class="tPad15">
			<form name="refreshFrm" method="get">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="page" value="">
			<table class="tbType1 listTb">
				<tr bgcolor="#FFFFFF">
					<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
					<td style="text-align:left;">
						<% SelectList "cate",CateCode %>
						<select name="isusingbox">
						<option value=""<% if isusing = "" then response.write " selected"%>>��뿩��</option>
						<option value="Y" <% if isusing = "Y" then response.write " selected"%>>Y</option>
						<option value="N" <% if isusing = "N" then response.write " selected"%>>N</option>
						</select>
						<select name="mdpick">
							<option value="">�����ȿ���</option>
							<option value="x" <% if mdpick = "x" then response.write " selected"%>>x</option>
							<option value="o" <% if mdpick = "o" then response.write " selected"%>>o</option>
						</select>
						<select name="limited">
							<option value="">limited ����</option>
							<option value="x" <% if limited = "x" then response.write " selected"%>>x</option>
							<option value="o" <% if limited = "o" then response.write " selected"%>>o</option>
						</select>
						<br><br>&nbsp;�귣��:
						<% drawSelectBoxDesignerwithName "ebrand", sBrand %>
						
					</td>
					<td style="text-align:left;">
						��ǰ �ڵ�:
						<textarea rows="2" cols="40" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
					</td>
					<td width="50" bgcolor="<%= adminColor("gray") %>">
						<input type="button" class="button_s" value="�˻�" onclick="refreshFrm.submit();">
					</td>
				</tr>
			</table>
			</form>
		</div>
		<!-- �˻� �� -->
		<!-- �׼� ���� -->
		<div class="tPad15">
			<form name="frmarr" method="post" action="">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="mode" value="">
			<table class="tbType1 listTb">
			<tr>
				<td style="text-align:left;">
					<input type="button" value="����Ⱥ귣��" onclick="fnspbrand();" class="button">
					<input type="button" value="�̹�������" onclick="contents_option();" class="button">
					<input type="button" value="Ű�������" onclick="keyword_option();" class="button">
					<input type="button" value="1+1 ����" onclick="pop1plus1Reg();" class="button">
					<input type="button" value="MDPICK��������" onclick="popMdpickSort();" class="button">		
					<!--<input type="button" value="�귣�� ���ͺ� ����" onclick="popBrandInterview();" class="button">-->
					���ù�ȣ�� 0���� �����ϸ� ��ǰ�� ������� �ʽ��ϴ�.

					<!--<% if mdpick = "o" then %>&nbsp;&nbsp;&nbsp;<input type="button" value="MD's Pick ���������ϱ�" onclick="popMainMDPickReg();" class="button"><% end if %>-->
					<!--<input type="button" value="���ΰ���" onclick="popMainReg();" class="button">-->
					<!--<input type="button" value="�̺�Ʈ����" onclick="popeventReg();" class="button">-->
				</td>
				<td align="right">
					<input type="button" value="������ǰ���" class="button" onClick="popRegItems();" />
					<input type="button" value="���ϻ�ǰ���" class="button" onClick="popRegNew();" />
				</td>
			</tr>
			</table>
			</form>
			<form name="frmreg" method="post" action="/admin/Diary2009/Lib/DiaryRegProc.asp">
				<input type="hidden" name="mode" value="mdpickreg">
				<input type="hidden" name="did" value="">
				<input type="hidden" name="mdpick" value="">
			</form>
			<form name="frmSortIsusing" method="post" action="/admin/diary2009/diary_list_sortIsusing_proc.asp" style="margin:0px;">
				<input type="hidden" name="sortnoarr" value="">
				<input type="hidden" name="isusingarr" value="">
				<input type="hidden" name="detailidxarr" value="">
				<input type="hidden" name="mode" value="sortisusingedit">
			</form>
		</div>
		<!-- �׼� �� -->
		<div class="tPad15">
		<% If C_ADMIN_AUTH Then %>
			<table class="a">
			<tr>
				<td>
					�ų� �������� �Ǹ� ��踦 ���� ������̺��� ��. ���̺� : [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
					--insert into [db_diary2010].[dbo].[diary_everyyear_for_statistic]<br>
					select ItemID, '2012', 'd' from [db_diary2010].[dbo].[tbl_DiaryMaster]<br>
					where isUsing = 'Y'<br>
					�۾��ڴ� ���� ���̾�� ������ ���� �� �ݵ�� �Է� ��. �⵵���� 2012~2013�����ϰ�� 2013.<br>
					���̾�� ���� 'd', ���ų������� ���� 'o'.<br>
					<b>�� 2014���̾���»�ǰ���� MD���� ��뿩�θ� ��� N���� �س��� ���� �� ��ǰ���� �� ���� ����.</b><br>
				</td>
			</tr>
			</table>
		<% End If %>
		</div>
		<div class="tPad15">
			<!-- ����Ʈ ���� -->
			<table class="tbType1 listTb">
			<form name="fitem" method="post" style="margin:0px;">
				<% IF oDiary.FResultCount>0 Then %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="17" style="text-align:left;">
						�˻���� : <b><%= oDiary.FTotalCount %></b>
						&nbsp;
						������ : <b><%= page %>/ <%= oDiary.FTotalPage %></b>
					</td>
				</tr>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td rowspan="2"> MDpick<br><input type="checkbox" name="chkA" onClick="jsChkAll();"><br><input class="button" type="button" id="btnEditSel" value="����" onClick="jsSortIsusing();"></td>
					<td rowspan="2"> ��ȣ</td>
					<td rowspan="2"> ���� </td>
					<td rowspan="2"> �̹��� </td>
					<td rowspan="2"> ��ǰ��ȣ </td>
					<td rowspan="2"> ��ǰ�� </td>
					<td rowspan="2"> ��ü���̵� </td>
					<td rowspan="2"> ���ù�ȣ<br>��뿩�� üũ<br><input type="checkbox" name="chkB" onClick="jsjunsiChkAll();"></td>
					<td>���ù�ȣ</td>
					<td> ��뿩��<br>
						<select name="selisusing" onchange="jsIsusingChg(this.value)" class="select">
							<option value="N">N</option>
							<option value="Y">Y</option>
						</select>
					</td>
					<td rowspan="2">�ǸŰ�</td>
					<td rowspan="2">����</td>
					<td rowspan="2"> ��౸�� </td>
					<td rowspan="2"> keyword </td>
					<td rowspan="2"> �������� </td>
					<td rowspan="2"> ���� </td>
					<td rowspan="2"> ������<br/>�̹��� </td>
				</tr>
				<tr bgcolor="<%= adminColor("tabletop") %>">
					<td colspan="2"><input class="button" type="button" id="btnEditSel" value="�������,��뿩�� ����" onClick="jsjunsiSortIsusing();"></td>
				</tr>
				<% For i =0 To  oDiary.FResultCount -1 %>
				<tr bgcolor="#FFFFFF">
					<td >
						<input type="checkbox" name="chkI" onClick="AnCheckClick(this);" value="<%= oDiary.FItemList(i).FDiaryID %>" <% IF oDiary.FItemList(i).Fmdpick="o" THEN %>checked<% END IF %>>
						<input type="hidden" name="mdpickchk" value="">
					</td>
					<td > <%= oDiary.FItemList(i).FDiaryID %> </td>
					<td ><% cateList "cate",oDiary.FItemList(i).FCateCode %> </td>
					<td >
						<img src="<%= db2html(oDiary.FItemList(i).FImageList) %>" width="40" height="40" border="0" style="cursor:pointer">
					</td>
					<td ><a href="<%= wwwUrl %>/shopping/category_prd.asp?itemid=<%= oDiary.FItemList(i).Fitemid %>&cate=<%=oDiary.FItemList(i).FCateCode%>" target="_blank"><%= oDiary.FItemList(i).Fitemid %> </a></td>
					<td > <%= oDiary.FItemList(i).fitemname %> </td>
					<td > <%= oDiary.FItemList(i).fmakerid %> </td>
					<td >
						<input type="checkbox" name="chkJ" onClick="AnCheckClick(this);" value="<%= oDiary.FItemList(i).FDiaryID %>">
						<input type="hidden" name="junsichk" value="">
					</td>
					<td>
						<input type="text" size="2" maxlength="2" name="sortNo" value="<%=oDiary.FItemList(i).Fsorting%>" class="text">
					</td>
					
					<td>
						<input type="hidden" value="<%=oDiary.FItemList(i).fisusing%>" name="orgisusing">
						<% drawSelectBoxUsingYN "isusing", oDiary.FItemList(i).fisusing %>
					</td>
					<td>
						<%
						Response.Write FormatNumber(oDiary.FItemList(i).Forgprice,0)
						'���ΰ�
						if oDiary.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>("&CLng((oDiary.FItemList(i).Forgprice-oDiary.FItemList(i).Fsailprice)/oDiary.FItemList(i).Forgprice*100) & "%��)" & FormatNumber(oDiary.FItemList(i).Fsailprice,0) & "</font>"
						end if
						'������
						if oDiary.FItemList(i).FitemCouponYn="Y" then
							Select Case oDiary.FItemList(i).FitemCouponType
								Case "1"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oDiary.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(oDiary.FItemList(i).GetCouponAssignPrice(),0) & "</font>"
							end Select
						end if
					%>
					</td><%'�ǸŰ�%>
					<td>
						<%
						Response.Write fnPercent(oDiary.FItemList(i).Forgsuplycash,oDiary.FItemList(i).Forgprice,1)
						'���ΰ�
						if oDiary.FItemList(i).Fsailyn="Y" then
							Response.Write "<br><font color=#F08050>" & fnPercent(oDiary.FItemList(i).Fsailsuplycash,oDiary.FItemList(i).Fsailprice,1) & "</font>"
						end if
						'������
						if oDiary.FItemList(i).FitemCouponYn="Y" then
							Select Case oDiary.FItemList(i).FitemCouponType
								Case "1"
									if oDiary.FItemList(i).Fcouponbuyprice=0 or isNull(oDiary.FItemList(i).Fcouponbuyprice) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(oDiary.FItemList(i).Fbuycash,oDiary.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(oDiary.FItemList(i).Fcouponbuyprice,oDiary.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									end if
								Case "2"
									if oDiary.FItemList(i).Fcouponbuyprice=0 or isNull(oDiary.FItemList(i).Fcouponbuyprice) then
										Response.Write "<br><font color=#5080F0>" & fnPercent(oDiary.FItemList(i).Fbuycash,oDiary.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									else
										Response.Write "<br><font color=#5080F0>" & fnPercent(oDiary.FItemList(i).Fcouponbuyprice,oDiary.FItemList(i).GetCouponAssignPrice(),1) & "</font>"
									end if
							end Select
						end if
						%>
					</td><%'����%>
					<td><%=fnColor(oDiary.FItemList(i).Fmwdiv,"mw")%><br/>
						<%
							If oDiary.FItemList(i).Fdeliverytype = "1" Then
								response.write "�ٹ�"
							ElseIf oDiary.FItemList(i).Fdeliverytype = "2" Then
								response.write "����"
							ElseIf oDiary.FItemList(i).Fdeliverytype = "4" Then
								response.write "�ٹ�"
							ElseIf oDiary.FItemList(i).Fdeliverytype = "9" Then
								response.write "����"
							ElseIf oDiary.FItemList(i).Fdeliverytype = "7" Then
								response.write "����"
							End If
						%>
					</td><%'��౸��%>	
					<td >
						<input type="button" class="button" value="<%=chkiif(oDiary.FItemList(i).Foptcount > 0,"����("&oDiary.FItemList(i).Foptcount&")","���")%>" onClick="detail_view('<%= oDiary.FItemList(i).FDiaryID %>');">
					</td>
					<td >
						<input type="button" class="button" value="<%=chkiif(oDiary.FItemList(i).Fnejicount > 0,"����("&oDiary.FItemList(i).Fnejicount&")","���")%>" onclick="javascript:popInfoReg('<%= oDiary.FItemList(i).FDiaryID %>');">
						<!--<input type="button" class="button" value="���" onclick="popInfoReg('<%= oDiary.FItemList(i).FDiaryID %>');">-->
					</td>
					<!--<td align="center"><input type="button" class="button" value="���" onclick="popContReg('<%= oDiary.FItemList(i).FDiaryID %>');"></td>-->
					<td >
						<input type="button" class="button" value="����" onclick="popRegModi('<%= oDiary.FItemList(i).FDiaryID %>');">
					</td>
					<td><input type="button" class="button" value="���/����" onclick="popPrevImg('<%= oDiary.FItemList(i).FDiaryID %>');"></td>
				</tr>
				<% Next %>
			<% else %>
					<tr bgcolor="#FFFFFF">
						<td colspan="3" class="page_link">[�˻������ �����ϴ�.]</td>
					</tr>
			<% End IF %>
				<tr bgcolor="#FFFFFF">
					<td colspan="17" align="center">
					<!-- ������ ���� -->
						<a href="?page=1&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/pprev_btn.gif" width="10" height="10" border="0"></a>
						<% if oDiary.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oDiary.StartScrollPage-1 %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/prev_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<% for i = 0 + oDiary.StartScrollPage to oDiary.StartScrollPage + oDiary.FScrollCount - 1 %>
							<% if (i > oDiary.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oDiary.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %>&nbsp;&nbsp;</b></font></span>
							<% else %>
							<a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" class="list_link"><font color="#000000"><%= i %>&nbsp;&nbsp;</font></a>
							<% end if %>
						<% next %>
						<% if oDiary.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>&isusingbox=<%=isusing%>&cate=<%=catecode%>">&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;</a></span>
						<% else %>
						&nbsp;<img src="http://fiximage.10x10.co.kr/web2007/common/next_btn.gif" width="10" height="10" border="0">&nbsp;
						<% end if %>
						<a href="?page=<%= oDiary.FTotalpage %>&isusingbox=<%=isusing%>&cate=<%=catecode%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2007/common/nnext_btn.gif" width="10" height="10" border="0"></a>
					<!-- ������ �� -->
					</td>
				</tr>
			</form>
			</table>
			<!-- ����Ʈ �� -->
		</div>
	</div>
</div>
<% Set oDiary = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->