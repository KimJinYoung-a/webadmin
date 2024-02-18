<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/itemcls_upche_2014.asp"--> 
<!-- #include virtual="/designer/lib/incPageFunction.asp"-->
<%
dim itemid, itemname, sellyn, limityn, dispCate,isfinish
dim clsItem 
dim arrList,intLoop
dim iTotCnt, iCurrpage, iPageSize
dim sSort

itemid  = RequestCheckVar(request("itemid"),500) 
itemname = RequestCheckVar(request("itemname"),32) 
sellyn  = RequestCheckVar(request("sellyn"),10) 
limityn = RequestCheckVar(request("limityn"),10)
dispCate = requestCheckvar(request("disp"),16)
isfinish = requestCheckvar(request("selisF"),1)
iCurrpage= requestCheckvar(request("iCP"),10)
sSort =  requestCheckVar(request("sS"),2)

iPageSize = 30
IF iCurrpage = "" THEN iCurrpage = 1
IF sSort = "" THEN sSort = "DD"	
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
set clsItem = new CUpCheItemEdit
	clsItem.FRectMakerId = session("ssBctID")
	clsItem.FRectItemId = itemid
	clsItem.FRectItemname = itemname
	clsItem.FRectSellYN	= sellYN
	clsItem.FRectLimitYN = limityn
	clsItem.FRectDispCate	= dispCate
	clsItem.FRectIsFinish	= isfinish
	clsItem.FRectSort = sSort
	clsItem.FCurrPage		= iCurrpage
	clsItem.FPageSize		= iPageSize
	arrList = clsItem.fnGetItemEditResultList
	iTotCnt	= clsItem.FTotCnt
set clsItem = nothing
%>
<style> 
	#dialog {display:none; position:absolute;left:100;top:100; z-index:9100;background:#efefef; padding:10px;width:650;}
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
	
//������û ���
function jsModCancel(itemid,itemname,sellcash){
	if(confirm("������û�� ����Ͻðڽ��ϱ�?")){
		document.frmMod.itemidarr.value = itemid;
		document.frmMod.olditemname.value = itemname;
		document.frmMod.oldsellcash.value = sellcash;
		document.frmMod.submit();
	}
}
</script>
<form name="frmMod" method="post" action="/lectureadmin/itemmaster/Upche_item_reqMod_Proc.asp">
	<input type="hidden" name="hidM" value="C">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="olditemname" value="">
	<input type="hidden" name="oldsellcash" value="">
</form>
<!-- ǥ ��ܹ� ����-->  
<form name="frmSearch" method=get>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sS" value="<%= ssort %>">
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
					 ����ī�װ�: <!-- #include virtual="/academy/comm/dispCateSelectBox.asp"-->
				</td> 
			</tr>
			</table>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
		</td>
	</tr> 
</table>
<br>  
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
 <tr bgcolor="#FFFFFF">
	 <td colspan="13" align="left">�˻�����: <%= formatnumber(iTotCnt,0)%> </td>
 </tr>
 <tr  align="center" bgcolor="<%= adminColor("tabletop") %>"> 
 	<td  onClick="javascript:jsSort('I','1');" style="cursor:hand;">��ǰ�ڵ� <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
 	<td>�̹���</td>
 	<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">��ǰ�� <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
  <td>�ǸŰ�</td>
 	<td>���ް�</td>
 	<td>���޸���</td>
 	<td>�Ǹ�</td>
 	<td>����</td>
 	<td onClick="javascript:jsSort('D','3');" style="cursor:hand;">������û�� <img src="/images/list_lineup<%IF sSort="DD" THEN%>_bot<%ELSEIF sSort="DA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
 	<td><select name="selisF" onChange="jsSearch();">
 			<option value="">��ü</option>
 			<option value="N" <%IF isfinish="N" THEN%>selected<%END IF%>>���δ��</option>
 			<option value="D" <%IF isfinish="D" THEN%>selected<%END IF%>>�ݷ���</option>
 			<option value="Y" <%IF isfinish="Y" THEN%>selected<%END IF%>>���ΰ�</option>
			</select>	
 	</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
	%>
<tr bgcolor="#FFFFFF" align="center"> 
	<td rowspan="2"><%=arrList(0,intLoop)%></td>
	<td rowspan="2"><img src="<%=imgFingers%>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(11,intLoop)%>"></td> 
	<td align="left"  <% IF  arrList(34,intLoop) = "N" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%>>
		<%IF arrList(34,intLoop) = "N" THEN%>
		<%=arrList(22,intLoop)%><br>
		<font color="Red">-><%=arrList(23,intLoop)%></font>
		<%ELSE%>
		<%=arrList(2,intLoop)%>
		<%END IF%> 
		&nbsp;&nbsp;<a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><font color="blue">Ȯ���ϱ�</font></a></td>
 	<td <%IF  arrList(34,intLoop) = "P" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%> align="right"> 
 		<%IF arrList(34,intLoop) = "P" THEN%>
 		<%=formatnumber(arrList(24,intLoop),0)%><br>
 		<font color="red">-><%=formatnumber(arrList(26,intLoop),0)%></font>
 		<%ELSE%>
		<%=formatnumber(arrList(15,intLoop),0)%> 
		<%END IF%>
		<%'���ΰ�
			if arrList(14,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>("&CLng((arrList(15,intLoop)-arrList(17,intLoop))/arrList(15,intLoop)*100) & "%��)" & FormatNumber(arrList(17,intLoop),0) & "</font>"
			end if
			'������
			if arrList(19,intLoop)="Y" then
				IF arrList(20,intLoop) = "1" or arrList(20,intLoop) ="2" THEN 
						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(21,intLoop),arrList(3,intLoop)),0) & "</font>"
				END IF 
			end if
		%> 
		</td>
	<td  	<%IF arrList(34,intLoop) = "P" THEN%>bgcolor="#DDFFDD"<%ELSE%>rowspan="2"<%END IF%> align="right"> 
		<%IF arrList(34,intLoop) = "P" THEN%>
 		<%=formatnumber(arrList(25,intLoop),0)%><br>
 		<font color="red">-><%=formatnumber(arrList(27,intLoop),0)%></font>
 		<%ELSE%>
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
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(16,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(21,intLoop),arrList(3,intLoop)),1) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & fnPercent(arrList(21,intLoop),GetCouponAssignPrice(arrList(19,intLoop),arrList(20,intLoop),arrList(21,intLoop),arrList(3,intLoop)),1) & "</font>"
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
		<%=fnGetReqStatus(arrList(30,intLoop))%><br>
		<%IF arrList(30,intLoop)="N" THEN%>
		 <a href="javascript:jsModCancel(<%=arrList(0,intLoop)%>,'<%=arrList(22,intLoop)%>','<%=arrList(24,intLoop)%>');"><font color="gray">[������û���]</font></a>
		<%ELSEIF arrList(30,intLoop)="D" THEN%> 
		<div><%=arrList(32,intLoop)%></div>
		<div><font color="red"><%=arrList(29,intLoop)%></font></div>
		<%ELSEIF arrList(30,intLoop)="Y" THEN%>
		<%=arrList(32,intLoop)%>
		<%END IF%>
	</td>
</tr> 
<tr bgcolor="#DDFFDD" height="30">
	<td <%IF arrList(34,intLoop) = "P" THEN%>colspan="2"<%END IF%>><font color="red">��������: <%=arrList(28,intLoop)%></font></td>
</tr> 
<%Next
ELSE
%>
<tr  bgcolor="#FFFFFF">
	<td colspan="10" align="center">��ϵ� ������ �����ϴ�.</td>
</tr>
<%END IF%>
</table>
</form>
<table width="100%" cellpadding="10" cellspacing="0">
	<tr>
		<td align="center"><%Call sbDisplayPaging("iCP",iCurrpage, iTotCnt,iPageSize, 10,menupos )%></td>
	</tr>
</table> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->