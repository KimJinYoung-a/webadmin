<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ��ǰ���
' History : 2008.04.04 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemgiftcls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
Dim clsGItem 
Dim gCode, acURL
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
Dim strparm
Dim clsGift
Dim sTitle, dSDay, dEDay, igStatus,igType,igR1,igR2,igkCode,igkType, igkCnt, igkLimit, igkName,sgDelivery
	
gCode = requestCheckVar(Request("gC"),10)
acURL = Server.HTMLEncode("/admin/shopmaster/gift/giftitemProc.asp?gC="&gCode)
strParm = request("strParm")
'=== �ڵ尪�� ���� ��� back
	IF gCode = "" THEN	
%>
		<script language="javascript">
		<!--
			alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
			history.back();
		//-->
		</script>
	<%	dbget.close()	:	response.End
	END IF	

'=== ����ǰ ���� 
 set clsGift = new CGift 
 	clsGift.FGCode = gCode
 	clsGift.fnGetGiftConts
 	sTitle		= clsGift.FGName
 	igType		= clsGift.FGType      
	igR1		= clsGift.FGRange1     
	igR2 		= clsGift.FGRange2    	
	igkCode		= clsGift.FGKindCode  
	igkType		= clsGift.FGKindType  
	igkCnt		= clsGift.FGKindCnt   
	igkLimit	= clsGift.FGKindlimit
 	dSDay		= clsGift.FSDate   	
	dEDay		= clsGift.FEDate    
	igStatus	= clsGift.FGStatus	
	igkName 	= clsGift.FGKindName
	sgDelivery = clsGift.FGDelivery
  set clsGift = nothing	
 	IF igkLimit = 0 THEN igkLimit = ""	
'=== �Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����  
	iCurrpage = Request("iC")	'���� ������ ��ȣ


	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
		
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

'=== ����ǰ ��� ��ǰ ����Ʈ 
 set clsGItem = new CGiftItem
 	clsGItem.FGCode = gCode
 	clsGItem.FCPage = iCurrpage
 	clsGItem.FPSize = iPageSize
 	arrList = clsGItem.fnGetItemConts
 	iTotCnt = clsGItem.FTotCnt
 
 
 iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
 
 '�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	Dim  arrgifttype,arrgiftstatus		
	arrgifttype 	= fnSetCommonCodeArr("gifttype",False)
	arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)
%>
<script language="javascript">
<!--
// ����ǰ �߰� �˾�
function addnewItem(){
		var popwin;
		popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?acURL=<%=acURL%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}

// ������ �̵�
function jsGoPage(iP){
		document.frmitem.iC.value = iP;		
		document.frmitem.submit();	
}

//��ü����
var ichk;
ichk = 1;
	
function jsChkAll(){			
	    var frm, blnChk;
		frm = document.frmitem;
		if(!frm.chkI) return;
		if ( ichk == 1 ){
			blnChk = true;
			ichk = 0;
		}else{
			blnChk = false;
			ichk = 1;
		}
		
 		for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}
//����
function jsDel(sType, iValue){	
		var frm;		
		var sValue;		
		frm = document.frmitem;
		sValue = "";
		
		if (sType ==0) {
			if(!frm.chkI) return;
			
			if (frm.chkI.length > 1){
			for (var i=0;i<frm.chkI.length;i++){
				if(frm.chkI[i].checked){
				   	if (sValue==""){
						sValue = frm.chkI[i].value;		
				   	}else{
						sValue =sValue+","+frm.chkI[i].value;		
				   	}	
				}
			}	
			}else{
				if(frm.chkI.checked){
					sValue = frm.chkI.value;
				}	
			}
		
			if (sValue == "") {
				alert('���� ��ǰ�� �����ϴ�.');
				return;
			}
			document.frmDel.itemidarr.value = sValue;
		}else{
			document.frmDel.itemidarr.value = iValue;
		}	
		 
		if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?")){		
			document.frmDel.submit();
		}
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a">
<form name="frmitem" method="get" action="giftItemReg.asp">
<input type="hidden" name="iC">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr>
		<td colspan="2">
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ�ڵ�</td>
			<td bgcolor="#FFFFFF"><%=gCode%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����ǰ��</td>
			<td bgcolor="#FFFFFF"><%=sTitle%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
			<td bgcolor="#FFFFFF" colspan="3"><%=dSDay%> ~ <%=dEDay%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF"><%=igkLimit%></td>		
		</tr>
		<tr>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">����</td>
			<td bgcolor="#FFFFFF"><%=fnGetCommCodeArrDesc(arrgiftstatus,igStatus)%></td>				
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF"><%=fnGetCommCodeArrDesc(arrgifttype,igType)%>&nbsp; <%=igR1%>�̻�~ <%=igR2%>�̸�</td>	
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� </td>
			<td bgcolor="#FFFFFF"><%=igkCnt%> <%IF igkType =2 THEN%>(1+1)<%END IF%></td>
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">���� </td>
			<td bgcolor="#FFFFFF"><%=igkName%>(<%=igkCode%>)
			<td align="center" bgcolor="<%= adminColor("tabletop") %>">��� </td>
			<td bgcolor="#FFFFFF"><%=fnSetDelivery(sgDelivery)%></td>
			
			
		</tr>
		</table>
		</td>
	</tr>
	<tr height="40" valign="bottom">
		<td align="left">
			<input type="button" value="���û���" onClick="jsDel(0,'');" class="button">
		</td>
		<td align="right">	
			<input type="button" value="����ǰ �߰�" onclick="addnewItem();" class="button">
		</td>
	</tr>
	<tr>
		<td colspan="2"> 
			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			    <tr bgcolor="#FFFFFF">
			   		<td colspan="15" align="left">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
			   	</tr>
			    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    				    	
			    	<td align="center">��ǰID</td>
					<td align="center">�̹���</td>
					<td align="center">�귣��</td>
					<td align="center">��ǰ��</td>
					<td align="center">�ǸŰ�</td>
					<td align="center">���԰�</td>
					<td align="center">���</td>	
					<td align="center">�Ǹſ���</td>	
					<td align="center">��뿩��</td>	
					<td align="center">��������</td>				    
			    	<td>ó��</td>
			    </tr>			  
			   <%IF isArray(arrList) THEN 
			    	For intLoop = 0 To UBound(arrList,2)
			   %>
			    <tr align="center" bgcolor="#FFFFFF">    
			    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    				    
			    	<td>
			    		<!-- 2007/05/05 ������ ���� -- ǰ�� ǥ�� -->			    		
			    		<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
			    		<% if clsGItem.IsSoldOut(arrList(12,intLoop),arrList(14,intLoop),arrList(18,intLoop),arrList(19,intLoop)) then %>
			    			<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
			    		<% end if %>
			    	</td>
			    	<td><% if (Not IsNull(arrList(10,intLoop)) ) and (arrList(10,intLoop)<>"") then %>
					 <img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(10,intLoop)%>">
					<%end if%>
			    	</td>    	
			    	<td><%=db2html(arrList(1,intLoop))%></td>
			    	<td align="left">&nbsp;<%=db2html(arrList(2,intLoop))%></td>
			    	<td><%
						Response.Write FormatNumber(arrList(5,intLoop),0)
						'���ΰ�
						if arrList(16,intLoop)="Y" then
							Response.Write "<br><font color=#F08050>(��)" & FormatNumber(arrList(7,intLoop),0) & "</font>"
						end if
						'������
						if arrList(20,intLoop)="Y" then
							Select Case arrList(21,intLoop)
								Case "1"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)*((100-arrList(22,intLoop))/100),0) & "</font>"
								Case "2"
									Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-arrList(22,intLoop),0) & "</font>"
							end Select
						end if
					%></td>
			    	<td><%
			Response.Write FormatNumber(arrList(6,intLoop),0)
			'���ΰ�
			if arrList(16,intLoop)="Y" then
				Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
			end if
			'������
			if arrList(20,intLoop)="Y" then
				if arrList(21,intLoop)="1" or arrList(21,intLoop)="2" then
					if arrList(23,intLoop)=0 or isNull(arrList(23,intLoop)) then
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(6,intLoop),0) & "</font>"
					else
						Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(23,intLoop),0) & "</font>"
					end if
				end if
			end if
				%></td>
			    	<td><%= fnColor(clsGItem.IsUpcheBeasong(arrList(13,intLoop)),"delivery")%></td>    	
			    	<td><%= fnColor(arrList(12,intLoop),"yn") %></td>
			    	<td><%= fnColor(arrList(17,intLoop),"yn") %></td>
			    	<td><%= fnColor(arrList(14,intLoop),"yn") %></td>    				    				    
			    	<td><input type="button" value="����" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td>	
			    </tr>   
			   <%	Next
			   	ELSE
			   %>
			   	<tr  align="center" bgcolor="#FFFFFF">
			   		<td colspan="12">��ϵ� ������ �����ϴ�.</td>
			   	</tr>	
			   <%END IF%>
			</table>
			<!-- ����¡ó�� -->
			<%		
			iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
			
			If (iCurrpage mod iPerCnt) = 0 Then																
				iEndPage = iCurrpage
			Else								
				iEndPage = iStartPage + (iPerCnt-1)
			End If	
			%>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">			      
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(iCurrpage) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong><%=ix%></strong></font></a>
					<%		else %>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><%=ix%></a>
					<%
							end if
						next
					%>
			    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
					<% else %>[next]<% end if %>
			        </td>
			        <td  width="50" align="right"><a href="giftList.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
			    </tr>				  
			 </table>
		</td>	    
	</tr>
		
</form>
</table> 
<%
set clsGItem = nothing
%>
<!-- ���û���--->
<form name="frmDel" method="post" action="giftItemProc.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="gC" value="<%=gCode%>">
<input type="hidden" name="itemidarr" value="">
</form>
<!-- ǥ �ϴܹ� ��-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->