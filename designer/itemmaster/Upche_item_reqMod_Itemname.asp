<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü��ۻ�ǰ��ǰ������û
' Hieditor : 2014.03.17 ������ ���� 
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_upche_2014.asp"-->  
<!-- #include virtual="/lib/function_item.asp"-->
<!-- #include virtual="/designer/lib/incPageFunction.asp"-->
<%
dim itemid, itemname, sellyn, limityn, dispCate
dim clsItem 
dim arrList,intLoop
dim iTotCnt, iCurrpage, iPageSize
dim sSort

itemid  = RequestCheckVar(request("itemid"),500) 
itemname = RequestCheckVar(request("itemname"),32) 
sellyn  = RequestCheckVar(request("sellyn"),10) 
limityn = RequestCheckVar(request("limityn"),10)
dispCate = requestCheckvar(request("disp"),16)
iCurrpage= requestCheckvar(request("iCP"),10)
sSort =  requestCheckVar(request("sS"),2)

iPageSize = 30
IF iCurrpage = "" THEN iCurrpage = 1
IF sSort = "" THEN sSort = "ID"	
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

'��ü��� ��ǰ����Ʈ ��������(�ٹ�����)	
set clsItem = new CItem
	clsItem.FRectMakerId = session("ssBctID")
	clsItem.FRectItemId = itemid
	clsItem.FRectItemname = itemname
	clsItem.FRectSellYN	= sellYN
	clsItem.FRectLimitYN = limityn
	clsItem.FRectDispCate	= dispCate
	clsItem.FRectSort = sSort
	clsItem.FCurrPage		= iCurrpage
	clsItem.FPageSize		= iPageSize
	arrList = clsItem.fnGetItemUpcheBaesongList
	iTotCnt	= clsItem.FTotCnt
set clsItem = nothing
%>
<style> 
	#dialog {display:none; position:absolute;left:100px;top:100px; z-index:9100;background:#efefef; padding:10px;width:650;}
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
	
//���û�ǰ ������û 
function jsChkSubmit(){
	var chkV = false;
	var itemcount = 0;
	var frm = document.frm;
	var len = 64;
	frm.itemidarr.value="";
 	frm.olditemname.value= "";
  frm.itemname.value = "";
  frm.etcStr.value=""
  frm.itemcount.value="";
  		
	if(frm.chkI){ 
		if (!frm.chkI.length){
	        if (frm.chkI.checked){
	        	if(!frm.sCIN.value){
	        		alert("������ ��ǰ���� �Է����ּ���");
	        		frm.sCIN.focus();
	        		return;
	        	}
	        	
	        	if (returnByteCount(frm.sCIN.value) > len)
				{
					alert("��ǰ���� �ѱ۱��� "+parseInt(len/2)+"��, �������� "+len+"�� �̳��� ���ּ���.");
					frm.sCIN.focus();
					return;
				}
				
	          	frm.itemidarr.value=frm.chkI.value;
		 		frm.olditemname.value= frm.sOIN.value;
		  	    frm.itemname.value = frm.sCIN.value;  
	        	chkV = true;
	        	itemcount = 1;
	        }
	   }else{ 
	   	  for (var i=0;i<frm.chkI.length;i++){
	            if (frm.chkI[i].checked){
	            	if(!frm.sCIN[i].value){
				    	alert("������ ��ǰ���� �Է����ּ���");
				        frm.sCIN[i].focus();
				        return;
				     } 
				     
					if (returnByteCount(frm.sCIN[i].value) > len){
					
						alert("��ǰ���� �ѱ۱��� "+parseInt(len/2)+"��, �������� "+len+"�� �̳��� ���ּ���.");
						frm.sCIN[i].focus();
						return;
					}
				   	
				    if( frm.itemidarr.value==""){ 
					      frm.itemidarr.value=frm.chkI[i].value;
						  frm.olditemname.value= frm.sOIN[i].value;
						  frm.itemname.value = frm.sCIN[i].value;  	
					}else{
						  frm.itemidarr.value=frm.itemidarr.value+"|,|"+frm.chkI[i].value;
						  frm.olditemname.value= frm.olditemname.value+"|,|"+frm.sOIN[i].value;
						  frm.itemname.value =frm.itemname.value+"|,|"+ frm.sCIN[i].value;  
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
	  frm.itemcount.value = itemcount;
	  var maskHeight = $(document).height();
		var maskWidth = $(document).width(); 
		$('#mask').css({'width':maskWidth,'height':maskHeight}); 
		$('#boxes').show();
		$('#mask').show(); 
		$("#dialog").show(); 

}

	$('#mask').click(function () {
		$('#boxes').hide();
		$('.window').hide();
		$('#dialog').hide(); 
	});
	
 
	function jsModSubmit(){
		if(!document.frmMS.sES.value){
			alert("��ǰ������������� �Է��� �ּ���.");
			document.frmMS.sES.focus();
			return;
		}
	  	document.frm.etcStr.value = document.frmMS.sES.value; 
		document.frm.submit();
	 
	}
	
	function jsCancel(){
			document.frm.itemidarr.value="";
 			document.frm.olditemname.value= "";
  		document.frm.itemname.value = "";
  		document.frm.etcStr.value=""
  		document.frm.itemcount.value="";
  	
  	 $( "#dialog" ).hide();
  	 $('#mask').hide();
  	 $('#boxes').hide();
	}
	
	//����Ʈ ����
function jsSort(sValue,i){ 
	 	document.frmSearch.sS.value= sValue;
	 	 
		   if (-1 < eval("document.frm.img"+i).src.indexOf("_alpha")){
	        document.frmSearch.sS.value= sValue+"D";  
	    }else if (-1 < eval("document.frm.img"+i).src.indexOf("_bot")){
	     		document.frmSearch.sS.value= sValue+"A";  
	    }else{
	       document.frmSearch.sS.value= sValue+"D";  
	    } 
		 document.frmSearch.submit();
	} 
	
	function jsChkClick(idx){
		var len=64; 
		if (!frm.chkI.length){
			document.frm.chkI.checked = true;
			if (returnByteCount(frm.sCIN.value) > len){
				alert("��ǰ����  �ѱ۱��� "+parseInt(len/2)+"��, �������� "+len+"�� �̳��� ���ּ���.");
				frm.sCIN.focus();
				return;
			}
		}else{
			document.frm.chkI[idx].checked = true;
			if (returnByteCount(frm.sCIN[idx].value) > len){ 
				alert("��ǰ���� �ѱ۱��� "+parseInt(len/2)+"��, �������� "+len+"�� �̳��� ���ּ���.");
				frm.sCIN[idx].focus();
				return;
			}
		}	
		
	} 
	 			
</script>
<!-- ǥ ��ܹ� ����-->    
<form name="frmSearch" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sS" value="<%=ssort %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left"> 
			 <table border="0" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %></td>
				<td> ����:<% drawSelectBoxLimitYN "limityn", limityn %> </td>
				<td> ��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="30"></td>
				<td> ��ǰ�ڵ�: </td>
				<Td rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea></td>
			</tr>
			<tr>
				<td colspan="4">	
					 ����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->
				</td> 
			</tr>
			</table>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
		</td>
	</tr> 
</table> 
</form> 

<table width="100%" border="0" class="a" >
<tr>
	<td align="left" style="padding-top:5px;">
		<input type="button" class="button" style="width:240px;background-color:#F8DFF0;" value="���û�ǰ�������û"   onClick="jsChkSubmit();"/> 
	</td> 
</tr>
</table>
<form name="frm" method="post" action="/designer/itemmaster/upche_item_reqMod_proc.asp">
<input type="hidden" name="hidM" value="N">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="olditemname" value="">
<input type="hidden" name="itemname" value="">
<input type="hidden" name="etcStr" value="">
<input type="hidden" name="itemcount" value=""> 
<input type="hidden" name="menupos" value="<%=menupos%>">
 
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
	 <td colspan="13" align="left">�˻�����: <%= formatnumber(iTotCnt,0)%> </td>
 </tr>
 <tr  align="center" bgcolor="<%= adminColor("tabletop") %>">
 	<td><input type="checkbox" name="chkAI" onClick="fnCheckAll(this.checked,frm.chkI);"></td>
 	<td  onClick="javascript:jsSort('I','1');" style="cursor:hand;">��ǰ�ڵ� <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
 	<td>�̹���</td>
 	<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">��ǰ�� <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
 	<td>�Ǹ�</td>
 	<td>����</td>
 	<td>�ǸŰ�</td>
 	<td>���ް�</td>
 	<td>���޸���</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>" onClick="AnCheckClick(this);"></td>
	<td><%=arrList(0,intLoop)%></td>
	<td><img src="<%=webImgUrl%>/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(11,intLoop)%>"></td> 
	<td align="left"><input type="hidden" name="sOIN" value="<%=arrList(2,intLoop)%>"><%=arrList(2,intLoop)%> &nbsp;&nbsp;<a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><font color="blue">Ȯ���ϱ�</font></a>
		<br>->����: <input type="text" name="sCIN" size="32" maxlength="64" onKeyUp="jsChkClick(<%=intLoop%>);">
		</td>
	<td><%=fnColor(arrList(5,intLoop),"yn")%></td>
	<td><%IF arrList(7,intLoop) ="Y" THEN%> 
      <%= fnColor(arrList(7,intLoop),"yn") %>
       <br>(<%= (arrList(8,intLoop) - arrList(9,intLoop)) %>)
      <% else %>	
       <%= fnColor(arrList(7,intLoop),"yn") %>
      <% end if %>
		
		</td>
<td align="right"><input type="hidden" name="mOSeP" value="<%=arrList(3,intLoop)%>">
		<%=formatnumber(arrList(15,intLoop),0)%> 
		<%'���ΰ�
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((arrList(15,intLoop)-arrList(17,intLoop))/arrList(15,intLoop)*100) & "%��)" & FormatNumber(arrList(17,intLoop),0) & "</font>"
			end if
			'������
			if arrList(19,intLoop)="Y" then
				IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),0) & "</font>"
				END IF 
			end if
		%>
		</td>
	<td align="right"><input type="hidden" name="mOSuP" value="<%=arrList(4,intLoop)%>">
		<%=formatnumber(arrList(16,intLoop),0)%> 
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
	<td>
		<%=fnPercent(arrList(16,intLoop),arrList(15,intLoop),1)%>
		<% '���ΰ�
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & fnPercent(arrList(18,intLoop),arrList(17,intLoop),1) & "</font>"
			end if
			'������
			if arrList(19,intLoop)="Y" then
					IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
			 			if  arrList(21,intLoop)=0 or isNull(arrList(21,intLoop)) then 
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(16,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(21,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(22,intLoop),arrList(3,intLoop)),1) & "</font>"
						end if
					END IF
			END IF			 
		%>
	</td>
</tr> 
<%Next
ELSE
%>
<tr  bgcolor="#FFFFFF">
	<td colspan="9" align="center">��ϵ� ������ �����ϴ�.</td>
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
<form name="frmMS" method="post"  onsubmit="return false;">
<div style="padding:10px;background-color:#FFFFFF"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ��ǰ�� ������û<hr></div> 
<table width="100%" border="0" class="a" cellpadding="5" cellspacing="5"  bgcolor="#FFFFFF">
	<tr >
		<td>��ǰ�� ������ <font color="red">���MD�� ���� �Ϸ� �� ����Ʈ�� �ݿ�</font>���� ���� ��Ź�帳�ϴ�.</td>
	</tr>
	<tr>
		<td>
		��ǰ�� ��������: <input type="text" name="sES" size="45" maxlength="64" value="">	
		</td>
	</tr>
	<tr>
		<td align="center">
			<input type="button" class="button" value="���" onClick="jsCancel();">
			<input type="button" class="button"  style="color:blue;" value="������û" onClick="jsModSubmit();">
		</td>
	</tr>
</table> 
</div> 
</form>


<!-- #include virtual="/lib/db/dbclose.asp" -->