<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ����ǰ ����
' History : 2010.09.27 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/gift/giftcls.asp"-->
<%
Dim eCode ,iSerachType,sSearchTxt,sGiftName,sBrand,  sDate,sSdate,sEdate,igStatus,sgDelivery
Dim clsGift, arrList, intLoop ,iTotCnt
Dim iPageSize, iCurrpage ,iDelCnt ,iStartPage, iEndPage, iTotalPage, ix,iPerCnt ,strParm
	eCode     		= requestCheckVar(Request("eC"),10)			'�̺�Ʈ �ڵ�
	iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
	sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
	sGiftName		= requestCheckVar(Request("sGN"),64)		'�˻� ����ǰ��
	sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
	sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
	sSdate     	= requestCheckVar(Request("iSD"),10)		'������
	sEdate     	= requestCheckVar(Request("iED"),10)		'������
	igStatus		= requestCheckVar(Request("giftstatus"),4)	'����ǰ ����
	sgDelivery		= requestCheckVar(Request("selDelivery"),1)	'�������
 
	iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
 
	IF iCurrpage = "" THEN	iCurrpage = 1
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	IF Cstr(eCode) = "0" THEN eCode = ""	
		
	IF (eCode <> "" AND sSearchTxt = "") THEN 
		iSerachType = "2"
		sSearchTxt = eCode
	ELSEIF 	(iSerachType="2" AND sSearchTxt <> "") THEN
		eCode = sSearchTxt
	END IF	

'�ڵ� ��ȿ�� �˻�(2008.08.04;������)
if sSearchTxt<>"" then
	if Not(isNumeric(sSearchTxt)) then
		if iSerachType="1" then
			Response.Write "<script language=javascript>alert('[" & sSearchTxt & "]��(��) ��ȿ�� ����ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
			dbget.close()	:	response.End
		else
			Response.Write "<script language=javascript>alert('[" & sSearchTxt & "]��(��) ��ȿ�� �̺�Ʈ�ڵ尡 �ƴմϴ�.');history.back();</script>"
			dbget.close()	:	response.End
		end if
	end if
end if

strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&giftstatus="&igStatus

set clsGift = new CGift
	clsGift.FECode = eCode
	clsGift.FSearchType = iSerachType    
	clsGift.FSearchTxt  = sSearchTxt     
	clsGift.FGiftName	= sGiftName  
	clsGift.FBrand		= sBrand     	
	clsGift.FDateType   = sDate     		
	clsGift.FSDate		= sSdate     	
	clsGift.FEDate		= sEdate     			
	clsGift.FGStatus	= igStatus
	clsGift.FGDelivery	= sgDelivery
	
 	clsGift.FCPage 		= iCurrpage
 	clsGift.FPSize 		= iPageSize
 	
	arrList = clsGift.fnGetGiftList	'�����͸�� ��������
	iTotCnt = clsGift.FTotCnt	'��ü ������  ��
set clsGift = nothing

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
Dim  arrgiftscope, arrgifttype,arrgiftstatus	

arrgiftscope 	= fnSetCommonCodeArr("giftscope",False)
arrgifttype 	= fnSetCommonCodeArr("gifttype",False)
arrgiftstatus 	= fnSetCommonCodeArr("giftstatus",False)	
%>

<script language="javascript">

	//�޷�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	//����
	function jsMod(gcode){
		location.href = "giftMod.asp?gC="+gcode+"&menupos=<%=menupos%>&<%=strParm%>";
	}
	
	//����¡ó��
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}
	
	//�̵�
	function jsGoURL(type,ival){
		if(type=="e"){		
			location.href = "/academy/event/event_modi.asp?evtId="+ival;
		}
	}
	
	//��ǰ������ �������̵�
	function jsItem(giftscope,gCode, eCode){
		//�̺�Ʈ��ϻ�ǰ, ���û�ǰ�ϋ� ��ǰ view, �׿� �������̵�
		if(giftscope == 2 || giftscope == 4 ){
			location.href = "/admin/eventmanage/event/eventitem_regist.asp?eC="+eCode+"&menupos=870";
		}else if(giftscope==5){
			location.href = "giftItemReg.asp?gC="+gCode+"&menupos=<%=menupos%>";
		}
	}

	function jsGoUrl(sUrl){
		self.location.href = sUrl;
	}
		
	
</script>

<!---- �˻� ---->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="get"  action="giftList.asp" onSubmit="return jsSearch(this,'E');">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="iC">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<select name="selType">
			<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>����ǰ�ڵ�</option>
			<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="10" maxlength="10">
		&nbsp;�귣��:
		<% drawSelectBoxLecturer "ebrand", sBrand %>
		&nbsp;����ǰ��:
		<input type="text" name="sGN" value="<%=sGiftName%>" maxlength="64" size="40">			
	</td>
	<td  rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frmSearch.submit();">
	</td>		
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		&nbsp;�Ⱓ:
		<select name="selDate">
			<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
			<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		</select>		
		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;">
		~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;">     
		&nbsp;����:		
		<%sbGetOptCommonCodeArr "giftstatus", igStatus, True,False,"onChange='javascript:document.frmSearch.submit();'"%>	
		&nbsp;���:		
		<select name="selDelivery" onChange="javascript:document.frmSearch.submit();">
			<option value="">��ü</option>
			<option value="Y" <%IF sgDelivery="Y" THEN%>selected<%END IF%>>��ü</option>
		<!--<option value="N" <%IF sgDelivery="N" THEN%>selected<%END IF%>>�ٹ�����</option>-->
		</select>
	</td>
</tr>	
</table>
<!---- /�˻� ---->

<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
    <tr height="40" valign="bottom">       
        <td align="left">
        	<input type="button" value="���ε��" class="button" onclick="javascript:location.href='giftreg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
        	<% if eCode <> "" then %><input type="button" value="�̺�Ʈ�������" onClick="jsGoUrl('/academy/event/event_list.asp?menupos=814');" class="button"><% end if %>
	    </td>
	    <td align="right"></td>        
	</tr>	
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="16">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>����ǰ�ڵ�</td>
	<td>�̺�Ʈ�ڵ�</br>(�׷�)</td>
	<td>����ǰ��</td>
	<td>�귣��</td>
	<td>�������</td>
	<td>��������</td>    	
	<td>�̻�</td>    	
	<td>�̸�</td>    	
	<td>����</td>
	<td>����</td>
	<td>������</td>
	<td>������</td>
	<td>����</td>
	<td>����</td>
	<td>���</td>
	<td>�����</td>
</tr>        
<%IF isArray(arrList) THEN
	For intLoop = 0 To UBound(arrList,2)
%> 
<% if arrList(17,intLoop) = "Y" then %>
<tr align="center" bgcolor="#FFFFFF">
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% end if %>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=arrList(0,intLoop)%></a></td>
	<td nowrap><%IF arrList(3,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(3,intLoop)%>)" title="�̺�Ʈ ��������"><%=arrList(3,intLoop)%></a><%IF arrList(4,intLoop) > 0 THEN%><br>(<%=arrList(4,intLoop)%>)<%END IF%><%END IF%></td>
	<td align="left">&nbsp;<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=db2html(arrList(1,intLoop))%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=db2html(arrList(5,intLoop))%></a></td>    	
	<td> <%IF (arrList(2,intLoop) = 2 or arrList(2,intLoop) = 4 or arrList(2,intLoop) = 5) then %>
		<a href="javascript:jsItem(<%=arrList(2,intLoop)%>,<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>)" title="��ϻ�ǰ ����"><%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%><br>(<%=arrList(20,intLoop)%>)</a>
		<%else%>
		<%=fnGetCommCodeArrDesc(arrgiftscope,arrList(2,intLoop))%>
		<%end if%>
		</td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=fnGetCommCodeArrDesc(arrgifttype,arrList(6,intLoop))%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=formatnumber(arrList(7,intLoop),0)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=formatnumber(arrList(8,intLoop),0)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=arrList(11,intLoop)%></a></td>
	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%IF arrList(9,intLoop) > 0 THEN%>[<%=arrList(9,intLoop)%>]<%=arrList(19,intLoop)%><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(22,intLoop) <> "" THEN %><%=arrList(22,intLoop)%><%END IF%>"><%=arrList(13,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="<%IF arrList(23,intLoop) <> "" THEN %><%=arrList(23,intLoop)%><%END IF%>"><%=arrList(14,intLoop)%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=fnGetCommCodeArrDesc(arrgiftstatus,arrList(15,intLoop))%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%IF arrList(12,intLoop) > 0 THEN%><%=arrList(12,intLoop)%><%END IF%></a></td>
		<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%IF arrList(21,intLoop)="Y" THEN%><font color="#F08050">��ü</font><%ELSE%><font color="#5080F0">�ٹ�����</font><%END IF%></a></td>
	<td nowrap><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="����ǰ ��������"><%=FormatDate(arrList(16,intLoop),"0000.00.00")%></a></td>    	
</tr>    	
<% Next
ELSE
%>
<tr>
	<td colspan="16" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
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
    </tr>    
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->