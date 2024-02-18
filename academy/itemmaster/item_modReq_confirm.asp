<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/itemcls_upche_2014.asp"-->  
<!-- #include virtual="/designer/lib/incPageFunction.asp"-->
<%
dim makerid, itemid, itemname, dstartdate, denddate, sreqtype, dispCate,isfinish
dim clsItem
dim arrList,intLoop
dim iTotCnt, iCurrpage, iPageSize
dim sSort, isRectSearch
dim onlyNotSet

makerid     = RequestCheckvar(request("makerid"),32)
itemid  	= RequestCheckVar(request("itemid"),500)
itemname 	= RequestCheckVar(request("itemname"),32)
dispCate 	= RequestCheckvar(request("disp"),16)
dstartdate  = RequestCheckVar(request("dSD"),10)
denddate  	= RequestCheckVar(request("dED"),10)
sreqtype 	= RequestCheckVar(request("selRT"),1)
isfinish 	= RequestCheckvar(request("selisF"),1)
iCurrpage	= RequestCheckvar(request("iCP"),10)
sSort 		= RequestCheckVar(request("sS"),2)
isRectSearch 	= RequestCheckVar(request("isRS"),1)
onlyNotSet 		= RequestCheckVar(request("onlyNotSet"),1)

iPageSize = 30
IF iCurrpage = "" THEN iCurrpage = 1
IF sSort = "" THEN sSort = "DD"
IF isRectSearch = "" THEN isfinish = "N"

'��ǰ�ڵ� ��ȿ���˻�
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
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

	if (onlyNotSet = "Y") then
		dispCate = ""
	end if

'��ü��� ��ǰ����Ʈ ��������(�ٹ�����)
set clsItem = new CItem
	clsItem.FRectMakerId 	= makerid
	clsItem.FRectItemId 	= itemid
	clsItem.FRectItemname 	= itemname

	if (onlyNotSet = "Y") then
		clsItem.FRectDispCate 	= "n"
	else
		clsItem.FRectDispCate 	= dispCate
	end if

	''clsItem.FRectDispCate	= dispCate
	clsItem.FRectStartDate 	= dStartDate
	clsItem.FRectEndDate 	= dEndDate
	clsItem.FRectReqType 	= sReqType
	clsItem.FRectIsFinish	= isfinish
	clsItem.FRectSortDiv 	= sSort
	clsItem.FCurrPage		= iCurrpage
	clsItem.FPageSize		= iPageSize

	arrList = clsItem.fnGetItemEditRequestList
	iTotCnt	= clsItem.FTotCnt
set clsItem = nothing
%>
<style>
	#dialog {display:none; position:absolute;left:100;top:100; z-index:9100;background:#efefef; padding:10px;width:400;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://webadmin.10x10.co.kr/images/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	//�˻�
	function jsSearch(){
			//��ǰ�ڵ� ����&���͸� �Է°����ϵ��� üũ-----------------------------
	var itemid = document.frmSearch.itemid.value;
	 itemid =  itemid.replace(",","\r");    //�޸��� �ٹٲ�ó��
		 for(i=0;i<itemid.length;i++){
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){
					alert("��ǰ�ڵ�� ���ڸ� �Է°����մϴ�.");
					return;
			}
		}
	//---------------------------------------------------------------------

	document.frmSearch.submit();
	}


	//����Ʈ ����
function jsSort(sValue,i){
	 	document.frmSearch.sS.value= sValue;

		   if (-1 < eval("document.frmSearch.img"+i).src.indexOf("_alpha")){
	        document.frmSearch.sS.value= sValue+"D";
	    }else if (-1 < eval("document.frmSearch.img"+i).src.indexOf("_bot")){
	     		document.frmSearch.sS.value= sValue+"A";
	    }else{
	       document.frmSearch.sS.value= sValue+"D";
	    }
		 document.frmSearch.submit();
	}

 function jsChkSubmit(sMode){
	var chkV = false;
	var itemcount = 0;
	var frm = document.frmItemList;

	if(frm.chkI){
	 if (!frm.chkI.length){
        if (frm.chkI.checked){
//        	if(sMode=="D"){
//	        	if(!frm.sRES.value){
//	        		alert("�ݷ������� �Է����ּ���");
//	        		frm.sRES.focus();
//	        		return;
//	        	}
//	        }
 
           if(frm.edittype.value=="P" && sMode=="A"){
        	// �����Էµ� ���ݺ��� ������ ������ 20%�̻� ���̰� ���� Ȯ�� �޽���
        	if(frm.sNewSC.value< parseInt(frm.sOldSC.value*0.8)) {
        		if(!confirm("\n\n\n\n��ǰ�ڵ�:"+frm.chkI.value+"�� �����û�� ������ ���� ���ݰ� �ſ� ���� ���̳��ϴ�(20%�̻�).\n\n������ ���� [ "+plusComma(frm.sOldSC.value)+" ]�� �� �����û��  ���� [ "+plusComma(frm.sNewSC.value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
        			return;
        		}
        	}
	         }
	         	document.frmF.itemid.value=frm.chkI.value;
	         	document.frmF.editidx.value=frm.editidx.value;
	         	document.frmF.edittype.value=frm.edittype.value;
		  	    document.frmF.itemname.value = frm.sNewIN.value;
		  	    document.frmF.sellcash.value = frm.sNewSC.value;
		  	    document.frmF.buycash.value = frm.sNewBC.value;
//		  	    document.frmF.rejectstr.value = frm.sRES.value;
	        	chkV = true;
	        	itemcount = 1;
        }
	   }else{
	   
	   	  for (var i=0;i<frm.chkI.length;i++){
	            if (frm.chkI[i].checked){
//	            	if(sMode=="D"){
//				        	if(!frm.sRES[i].value){
//				        		alert("�ݷ������� �Է����ּ���");
//				        		frm.sRES[i].focus();
//				        		return;
//				        	}
//			
                if(frm.edittype[i].value=="P" && sMode=="A"){	       
                    // �����Էµ� ���ݺ��� ������ ������ 20%�̻� ���̰� ���� Ȯ�� �޽���
                	if(frm.sNewSC[i].value< parseInt(frm.sOldSC[i].value*0.8)) {
                		if(!confirm("\n\n\n\n��ǰ�ڵ�:"+frm.chkI[i].value+"�� �����û�� ������ ���� ���ݰ� �ſ� ���� ���̳��ϴ�(20%�̻�).\n\n������ ���� [ "+plusComma(frm.sOldSC[i].value)+" ]�� �� �����û��  ���� [ "+plusComma(frm.sNewSC[i].value)+" ]��\n\n\n�Է��Ͻ� ������ ��Ȯ�մϱ�?\n\n\n\n")) {
                			return;
                		}
                	}
                 }	
				        	if( document.frmF.itemid.value==""){
						        document.frmF.itemid.value=frm.chkI[i].value;
						        document.frmF.editidx.value=frm.editidx[i].value;
						        document.frmF.edittype.value=frm.edittype[i].value;
						  	    document.frmF.itemname.value = frm.sNewIN[i].value;
						  	    document.frmF.sellcash.value = frm.sNewSC[i].value;
						  	    document.frmF.buycash.value = frm.sNewBC[i].value;
						  	   // document.frmF.rejectstr.value = frm.sRES[i].value;

						  	   }else{
						  	   	document.frmF.itemid.value=document.frmF.itemid.value+","+frm.chkI[i].value;
						  	   	document.frmF.editidx.value=document.frmF.editidx.value+","+frm.editidx[i].value;
						  	   	document.frmF.edittype.value=document.frmF.edittype.value+","+frm.edittype[i].value;
						  	    document.frmF.itemname.value = document.frmF.itemname.value+","+frm.sNewIN[i].value;
						  	    document.frmF.sellcash.value =document.frmF.sellcash.value+","+frm.sNewSC[i].value;
						  	    document.frmF.buycash.value = document.frmF.buycash.value+","+frm.sNewBC[i].value;
						  	   // document.frmF.rejectstr.value = document.frmF.rejectstr.value+","+frm.sRES[i].value;
						  	  }
				        itemcount = itemcount + 1;
	            	chkV = true;
	            }
	      }
	  }

	}

	 if (!chkV){
	  	alert("���õ� ��ǰ�� �����ϴ�.");
			return;
	  }

	 document.frmF.hidM.value =sMode;
	 document.frmF.itemcount.value = itemcount;

	 if(sMode=="D"){
	 	var maskHeight = $(document).height();
		var maskWidth = $(document).width();
		$('#mask').css({'width':maskWidth,'height':maskHeight});
		$('#boxes').show();
		$('#mask').show();
		$("#dialog").show();
	 }else{
		document.frmF.submit();
	 }
}

	$('#mask').click(function () {
		$('#boxes').hide();
		$('.window').hide();
		$('#dialog').hide();
	});

//���δ����·� ����
function jsChangeStatus(itemid,editidx){ 
	if(confirm("���δ�� ���·� �����Ͻðڽ��ϱ�?")){
	document.frmR.itemid.value = itemid;
	document.frmR.editidx.value = editidx; 
	document.frmR.submit();
}
}



function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function jsRejectSubmit(){
	if(!document.frmRej.sRES.value){
		alert("�ݷ������� �Է����ּ���");
		return;
	}

	document.frmF.rejectstr.value = document.frmRej.sRES.value;
	document.frmF.submit();
}

function jsCancel(){
	document.frmF.itemid.value	= "";
 	document.frmF.itemname.value= "";
	document.frmF.sellcash.value= "";
	document.frmF.buycash.value = "";
	document.frmF.hidM.value 	="";
	document.frmF.itemcount.value= "";

  	 $("#dialog").hide();
  	 $("#mask").hide();
  	 $("#boxes").hide();
	}
</script>
 <form name="frmR" method="post" action="/academy/itemmaster/item_modReq_confirm_proc.asp">
 	<input type="hidden" name="hidM" value="C">
 	<input type="hidden" name="itemid" value="">
 	<input type="hidden" name="editidx" value=""> 
 	<input type="hidden" name="itemname" value="">
 	<input type="hidden" name="sellcash" value="">
 	<input type="hidden" name="buycash" value="">
 	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rS" value="<%= ssort %>">
	<input type="hidden" name="rmakerid" value="<%= makerid %>">
	<input type="hidden" name="ritemname" value="<%= itemname %>">
	<input type="hidden" name="ritemid" value="<%= itemid %>">
	<input type="hidden" name="rdispCate" value="<%= dispCate %>">
	<input type="hidden" name="rSD" value="<%= dStartDate %>">
	<input type="hidden" name="rED" value="<%= dEndDate %>">
	<input type="hidden" name="rRT" value="<%= sreqtype %>">
</form>
 <form name="frmF" method="post" action="/academy/itemmaster/item_modReq_confirm_proc.asp">
 	<input type="hidden" name="hidM" value="">
 	<input type="hidden" name="itemid" value="">
 	<input type="hidden" name="editidx" value="">
 	<input type="hidden" name="edittype" value="">
 	<input type="hidden" name="itemname" value="">
 	<input type="hidden" name="sellcash" value="">
 	<input type="hidden" name="buycash" value="">
 	<input type="hidden" name="rejectstr" value="">
 	<input type="hidden" name="itemcount" value="">
 	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rS" value="<%= ssort %>">
	<input type="hidden" name="rmakerid" value="<%= makerid %>">
	<input type="hidden" name="ritemname" value="<%= itemname %>">
	<input type="hidden" name="ritemid" value="<%= itemid %>">
	<input type="hidden" name="rdispCate" value="<%= dispCate %>">
	<input type="hidden" name="rSD" value="<%= dStartDate %>">
	<input type="hidden" name="rED" value="<%= dEndDate %>">
	<input type="hidden" name="rRT" value="<%= sreqtype %>">
</form>
<!-- ǥ ��ܹ� ����-->
<form name="frmSearch" method="get" action="item_modReq_confirm.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sS" value="<%= ssort %>">
<input type="hidden" name="isRS" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			 <table border="0" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td> �귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
				<td> ��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30"></td>
				<td> ��ǰ�ڵ�: </td>
				<Td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
			</tr>
			<tr>
				<td colspan="4">
					 ����ī�װ�: <!-- #include virtual="/academy/comm/dispCateSelectBox.asp"-->
				</td>
			</tr>
			<tr>
				<td>��û�Ⱓ:
						<input type="text" size="10" name="dSD" value="<%=dStartDate%>" onClick="jsPopCalendar('frmSearch','dSD');" style="cursor:hand;">
						<img src="/images/admin_calendar.png" alt="�޷����� �˻�" onClick="jsPopCalendar('frmSearch','dSD');"/>
					~ <input type="text" size="10" name="dED" value="<%=dEndDate%>" onClick="jsPopCalendar('frmSearch','dED');"  style="cursor:hand;">
						<img src="/images/admin_calendar.png" alt="�޷����� �˻�" onClick="jsPopCalendar('frmSearch','dED');"/>
				</td>
				<td>
					��û����: <select name="selRT">
						<option value="">��ü</option>
						<option value="N" <%IF sreqtype="N" THEN%>selected<%END IF%>>��ǰ��</option>
						<option value="P" <%IF sreqtype="P" THEN%>selected<%END IF%>>��ǰ����</option>
					</select>
					&nbsp;
					����:
					<select name="selisF" onChange="jsSearch();">
						<option value="">��ü</option>
						<option value="N" <%IF isfinish="N" THEN%>selected<%END IF%>>���δ��</option>
						<option value="D" <%IF isfinish="D" THEN%>selected<%END IF%>>�ݷ���</option>
						<option value="Y" <%IF isfinish="Y" THEN%>selected<%END IF%>>���ΰ�</option>
					</select>
					</td>
			</tr>
			<tr>
				<td colspan="2"><input type="checkbox" name="onlyNotSet" value="Y" <% if (onlyNotSet = "Y") then %>checked<% end if %> > ����ī�װ� ������ ��ǰ��</td>
			</tr>
			</table>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
		</td>
	</tr>
</table>
</form>
<form name="frmItemList" method="get" action="">
<table width="100%" border="0" class="a" >
<tr>
	<td align="left" style="padding-top:5px;">
		<input type="button" class="button" style="width:80px;background-color:#F8DFF0;" value="����"   onClick="jsChkSubmit('A');"/>
		<input type="button" class="button" style="width:80px;background-color:#F8DFF0;" value="�ݷ�"   onClick="jsChkSubmit('D');"/>
	</td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
	 <td colspan="13" align="left">�˻�����: <%= formatnumber(iTotCnt,0)%> </td>
 </tr>
 <tr  align="center" bgcolor="<%= adminColor("tabletop") %>">
 	<td><input type="checkbox" name="chkAI" onClick="CkeckAll(this);"></td>
 	<td  onClick="javascript:jsSort('I','1');" style="cursor:hand;">��ǰ�ڵ� <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
 	<td>�̹���</td>
 	<td onClick="javascript:jsSort('B','3');" style="cursor:hand;">�귣��ID <img src="/images/list_lineup<%IF sSort="BD" THEN%>_bot<%ELSEIF sSort="BA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
 	<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">��ǰ�� <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
  <td>�ǸŰ�</td>
 	<td>���ް�</td>
 	<td>����</td>
 	<td>�Ǹ�</td>
 	<td>����</td>
 	<td onClick="javascript:jsSort('D','4');" style="cursor:hand;">������û�� <img src="/images/list_lineup<%IF sSort="DD" THEN%>_bot<%ELSEIF sSort="DA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
 	<td>
 	</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%>
	<input type="hidden" name="editidx" value="<%=arrList(34,intLoop)%>">
	<input type="hidden" name="edittype" value="<%=arrList(35,intLoop)%>">
<tr bgcolor="#FFFFFF" align="center">
	<td rowspan="2"><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>" onClick="AnCheckClick(this);" <%IF arrList(30,intLoop)<>"N" THEN%>disabled<%END IF%>></td>
	<td rowspan="2"><%=arrList(0,intLoop)%></td>
	<td rowspan="2"><img src="<%=imgFingers%>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(11,intLoop)%>"></td>
	<td rowspan="2"><%=arrList(1,intLoop)%></td>
	<td align="left"  <% IF arrList(35,intLoop) = "N" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%>>
		<%IF arrList(35,intLoop) = "N" THEN%>
		<input type="hidden" name="sNewIN" value="<%=replace(arrList(23,intLoop),"""","&quot;")%>">
		<%=arrList(22,intLoop)%><br>
		<font color="Red">-><%=arrList(23,intLoop)%></font>
		<%ELSE%>
		<input type="hidden" name="sNewIN" value="">
		<%=arrList(2,intLoop)%>
		<%END IF%>
		&nbsp;&nbsp;<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><font color="blue">Ȯ���ϱ�</font></a></td>

 	<td <%IF arrList(35,intLoop) = "P" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%> align="right">
 		<%IF arrList(35,intLoop) = "P" THEN%>
 			<input type="hidden" name="sOldSC" value="<%=arrList(24,intLoop)%>">
 			<input type="hidden" name="sNewSC" value="<%=arrList(26,intLoop)%>">
 		<%=formatnumber(arrList(24,intLoop),0)%><br>
 		<font color="red">-><%=formatnumber(arrList(26,intLoop),0)%></font>
 		<%ELSE%>
 		<input type="hidden" name="sNewSC" value="">
 		<input type="hidden" name="sOldSC" value="">
		<%=formatnumber(arrList(15,intLoop),0)%>
		<%END IF%>
		<%'���ΰ�
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((arrList(15,intLoop)-arrList(17,intLoop))/arrList(15,intLoop)*100) & "%��)" & FormatNumber(arrList(17,intLoop),0) & "</font>"
			end if
			'������
			if arrList(19,intLoop)="Y" then
				IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN
					Response.Write "<br><font color=#5080F0>(��)"	 & FormatNumber(GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(33,intLoop),arrList(3,intLoop)),0) & "</font>"
				END IF
			end if
		%>
		</td>
	<td  	<%IF arrList(35,intLoop) = "P" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%> align="right">
		<%IF arrList(35,intLoop) = "P" THEN%>
		<input type="hidden" name="sNewBC" value="<%=arrList(27,intLoop)%>">
 		<%=formatnumber(arrList(25,intLoop),0)%><br>
 		<font color="red">-><%=formatnumber(arrList(27,intLoop),0)%></font>
 		<%ELSE%>
 		<input type="hidden" name="sNewBC" value="">
		<%=formatnumber(arrList(16,intLoop),0)%>
		<%END IF%>
	 <%	'����
	 		if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(18,intLoop),0) & "</font>"
			end if
			'������
		if arrList(19,intLoop)="Y" then
			IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN
					if  arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(16,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(21,intLoop),0) & "</font>"
					end if
			END IF
		END IF
		%>
	</td>
	<td  rowspan="2">
		<%=fnPercent(arrList(16,intLoop),arrList(15,intLoop),1)%>
		<% '���ΰ�
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(arrList(18,intLoop),arrList(17,intLoop),1) & "</font>"
			end if
			'������
			if arrList(19,intLoop)="Y" then
					IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN
			 			if  arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(16,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(33,intLoop),arrList(3,intLoop)),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(21,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(33,intLoop),arrList(3,intLoop)),1) & "</font>"
						end if
					END IF
			END IF
		%>
	</td>
		<td rowspan="2"><%=fnColor(arrList(5,intLoop),"yn")%></td>
	<td rowspan="2"><%IF arrList(7,intLoop) ="Y" THEN%>
      <%= fnColor(arrList(7,intLoop),"yn") %>
       <br>(<%= (arrList(8,intLoop) - arrList(9,intLoop)) %>)
      <% else %>
       <%= fnColor(arrList(7,intLoop),"yn") %>
      <% end if %>

		</td>
	<td rowspan="2"><%=arrList(31,intLoop)%></td>
	<td rowspan="2">
		<div><%=fnGetReqStatus(arrList(30,intLoop))%></div>
		<%IF arrList(30,intLoop)="D" THEN%>
		<div><%=arrList(32,intLoop)%></div>
		<div><font color="red"><%=arrList(29,intLoop)%></font></div>
	<a href="javascript:jsChangeStatus('<%=arrList(0,intLoop)%>','<%=arrList(34,intLoop)%>');"><font color="gray">[���δ�⺯��]</font></a>
		<%ELSEIF arrList(30,intLoop)="Y" THEN%>
		<%=arrList(32,intLoop)%>
		<%END IF%>
	</td>
</tr>
<tr bgcolor="#DDFFDD" height="30">
	<td <%IF arrList(35,intLoop) = "P" THEN%>colspan="2"<%END IF%>><font color="red">��������: <%=arrList(28,intLoop)%></font></td>
</tr>
<%Next
ELSE
%>
<tr  bgcolor="#FFFFFF">
	<td colspan="12" align="center">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>
</form>
<table width="100%" cellpadding="10" cellspacing="0">
	<tr>
		<td align="center"><%Call sbDisplayPaging("iCP",iCurrpage, iTotCnt,iPageSize, 10,menupos )%></td>
	</tr>
</table>
<div id="boxes">
<div id="mask"></div>
<div id="dialog">

	<div style="padding:10px;background-color:#FFFFFF"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �ݷ�����<hr></div>
<form name="frmRej" method="post">
<table width="100%" border="0" class="a" cellpadding="5" cellspacing="5"  bgcolor="#FFFFFF">
	<tr >
		<td align="center"><input type="text" name="sRES" size="45" maxlength="64" value="">	</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="button" value="���" onClick="jsCancel();">
			<input type="button" class="button"  style="color:blue;" value="�ݷ�" onClick="jsRejectSubmit();">
		</td>
	</tr>
</table>
</form>
</div>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->