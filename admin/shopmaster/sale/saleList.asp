<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ����
' History : 2008.04.07 ������ ����
'			2021.07.13 �ѿ�� ����(�˻����� ���� ����)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
'Call fnSetEventCommonCode '�����ڵ� ���ø����̼� ������ ����

 Dim eCode
 Dim clsSale, arrList, intLoop
 Dim iTotCnt
 Dim iPageSize, iCurrpage ,iDelCnt
 Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
 Dim iSerachType,sSearchTxt,sBrand,  sDate,sSdate,sEdate,isStatus
 Dim isaleType
 Dim strParm
  
 eCode     		= requestCheckVar(Request("eC"),10)			'�̺�Ʈ �ڵ�
 iSerachType    = requestCheckVar(Request("selType"),4)		'�˻�����
 sSearchTxt     = requestCheckVar(Request("sTxt"),10)		'�˻���
 sBrand     	= requestCheckVar(Request("ebrand"),32)		'�귣��
 sDate     		= requestCheckVar(Request("selDate"),1)		'�˻��� ����
 sSdate     	= requestCheckVar(Request("iSD"),10)		'������
 sEdate     	= requestCheckVar(Request("iED"),10)		'������
 isStatus		= requestCheckVar(Request("salestatus"),4)	'���� ����
 isaleType      = requestCheckVar(Request("selST"),1)       '����Ÿ��
 arrList = ""
 iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
 if iSerachType = "" then iSerachType = 3
 	if iSerachType="1" or iSerachType="2" then
 		'�˻��κ��� ��ȣ�� �޾ƾߵȴٸ� ���ڸ� ����
 		sSearchTxt = getNumeric(sSearchTxt)
 	end if
 
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	IF Cstr(eCode) = "0" THEN eCode = ""	
	IF (eCode <> "" AND sSearchTxt = "") THEN 
		iSerachType = 2
		sSearchTxt = eCode
	END IF
				
    strParm =  "iC="&iCurrpage&"&eC="&eCode&"&selType="&iSerachType&"&sTxt="&sSearchTxt&"&ebrand="&sBrand&"&selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&sstatus="&isStatus    
	set clsSale = new CSale
		clsSale.FECode = eCode
		clsSale.FSearchType = iSerachType    
 		clsSale.FSearchTxt  = sSearchTxt     
 		clsSale.FBrand		= sBrand     	
 		clsSale.FDateType   = sDate     		
 		clsSale.FSDate		= sSdate     	
 		clsSale.FEDate		= sEdate     			
 		clsSale.FSStatus	= isStatus
	 	clsSale.FCPage 		= iCurrpage
	 	clsSale.FPSize 		= iPageSize
	 	clsSale.FSType      = isaleType
	 	
		arrList = clsSale.fnGetSaleList	'�����͸�� ��������
 		iTotCnt = clsSale.FTotCnt	'��ü ������  ��
 	set clsSale = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��	
	
	Dim arrsalemargin, arrsalestatus
	'�����ڵ� �� �迭�� �Ѳ����� ������ �� �� �����ֱ�
	arrsalemargin = fnSetCommonCodeArr("salemargin",False)
	arrsalestatus= fnSetCommonCodeArr("salestatus",False)	
%>
<script type="text/javascript">

	//�޷�
	function jsPopCal(sName){
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
	
	//����
	function jsMod(scode){
		location.href = "saleReg.asp?sC="+scode+"&menupos=<%=menupos%>&<%=strParm%>";
	}
	
	//����¡ó��
		function jsGoPage(iP){
		document.frmSearch.iC.value = iP;
		document.frmSearch.submit();
	}
	
	//�̵�
	function jsGoURL(type,ival){
		if(type=="e"){		
			location.href = "/admin/eventmanage/event/event_modify.asp?eC="+ival;
		}else if(type=="i"){
			location.href = "saleReg.asp?sC="+ival+"&menupos=<%=menupos%>";
		}
	}
	
	//���� �ٷ� ����ó��
 	function jsSetRealSale(sCode, chkState){  
 		if(chkState !=1){
 			alert("�������̰� ���糯¥�� �̺�Ʈ �Ⱓ���϶��� �ǽð� ó�� �����մϴ�.");
 			return;
 		}
 		
 		if(confirm("��ϵ� ����ǰ�� ���� ����� �������� �ٷ� ����˴ϴ�. ó���Ͻðڽ��ϱ�?")){
 			document.frmReal.sC.value = sCode;
 			document.frmReal.submit();
 		}
 	}

	function frmsubmit(page){
		if (frmSearch.selDate.value!=''){
			if (frmSearch.iSD.value!=''){
				if (frmSearch.iED.value==''){
					alert('�������� �Է��� �ּ���.');
					return;
				}
			}
			if (frmSearch.iED.value!=''){
				if (frmSearch.iSD.value==''){
					alert('�������� �Է��� �ּ���.');
					return;
				}
			}
		}
		frmSearch.iC.value=page;
		frmSearch.submit()
	}

</script>
<!---- �˻� ---->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">	
<form name="frmReal" method="post" action="saleItemProc.asp?<%=strParm%>">
<input type="hidden" name="sC">
<input type="hidden" name="mode" value="P">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
	<form name="frmSearch" method="get"  action="saleList.asp" onSubmit="return jsSearch(this,'E');">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="iC">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		<select name="selType" class="select">
		<option value="1" <%IF Cstr(iSerachType) = "1" THEN%>selected<%END IF%>>�����ڵ�</option>
		<option value="2" <%IF Cstr(iSerachType) = "2" THEN%>selected<%END IF%>>�̺�Ʈ�ڵ�</option>
		<option value="3" <%IF Cstr(iSerachType) = "3" THEN%>selected<%END IF%>>���θ�</option>
		</select>
		<input type="text" name="sTxt" value="<%=sSearchTxt%>" size="20" maxlength="64" class="input">		
		&nbsp;�Ⱓ:
		<select name="selDate" class="select">
		<option value="S" <%if Cstr(sDate) = "S" THEN %>selected<%END IF%>>������ ����</option>
		<option value="E" <%if Cstr(sDate) = "E" THEN %>selected<%END IF%>>������ ����</option>
		</select>		
		<input type="text" size="10" name="iSD" value="<%=sSdate%>" onClick="jsPopCal('iSD');" style="cursor:hand;" class="input">
		~ <input type="text" size="10" name="iED" value="<%=sEdate%>" onClick="jsPopCal('iED');"  style="cursor:hand;" class="input">      
		&nbsp;����:
		<%sbGetOptCommonCodeArr  "salestatus", isStatus, True, False,"onChange='javascript:document.frmSearch.submit();'"%>		
		&nbsp����Ÿ��:
		<select name="selST" class="select">
		<option value="">--��ü--</option>    
		<option value="1" <%if isaleType= "1" THEN %>selected<%END IF%>>�Ϲ� ����</option>
		<option value="2" <%if isaleType = "2" THEN %>selected<%END IF%>>Ÿ�� Ư��</option>
		</select>		
		</td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('1');">
		</td>
	</tr>	
</table>
<!---- /�˻� ---->
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a"  >	
    <tr height="40" valign="bottom">       
        <td align="left">
        	<input type="button" value="���ε��" class="button" onclick="javascript:location.href='saleReg.asp?menupos=<%=menupos%>&eC=<%=eCode%>';" >
	    </td>
	    <td align="right"></td>        
	</tr>	
</table>
<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="13">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	</tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>�����ڵ�</td>
    	<td>�̺�Ʈ�ڵ�</br>(�׷��ڵ�)</td>
    	<td>���θ�</td>    	    	
    	<td>������</td>
    	<td>��������</td>
    	<td>������</td>
    	<td>������</td>
    	<td>����</td>    	
    	<td>����Ÿ��</td>   
    	<td>��ǰ����<br>����ð�</td>
    	<td colspan="2">ó��</td>
    	<td>�����</td>
    </tr>        
    <% Dim chkState  
    IF isArray(arrList) THEN 
    	For intLoop = 0 To UBound(arrList,2)  
    	chkState = 0  	
    	'����: ����, �����û )�Ⱓ: �����ϱ��� �Ⱓ��
    	if (arrList(8,intLoop) = 6 or arrList(8,intLoop) = 7 or arrList(8,intLoop) = 9) and datediff("d",arrList(6,intLoop),date()) >=0 and datediff("d",arrList(7,intLoop),date()) <=0 then
    		chkState = 1    	
    	end if	
    %> 
    <tr align="center" bgcolor="#FFFFFF">    
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=arrList(0,intLoop)%></a></td>
    	<td><%IF arrList(4,intLoop) > 0 THEN%><a href="javascript:jsGoURL('e',<%=arrList(4,intLoop)%>)" title="�̺�Ʈ ��������"><%=arrList(4,intLoop)%></a><%IF arrList(5,intLoop) > 0 THEN%><br>(<%=arrList(5,intLoop)%>)<%END IF%><%END IF%></td>
    	<td align="left">&nbsp;<a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=db2html(ReplaceBracket(arrList(1,intLoop)))%></a></td>    	
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=arrList(2,intLoop)%>%</a></td>    
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=fnGetCommCodeArrDesc(arrsalemargin,arrList(3,intLoop))%></a></td>     
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=arrList(6,intLoop)%></a></td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=arrList(7,intLoop)%></a></td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%IF arrList(8,intLoop) = 6 THEN%><font color="blue"><%END IF%><%=fnGetCommCodeArrDesc(arrsalestatus,arrList(8,intLoop))%></a></td>
    	<td><%if arrList(17,intLoop)=2 then%><font color="orange">Ÿ��Ư��</font><%else%>�Ϲ�����<%end if%></td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������">
    		<%IF arrList(8,intLoop) = 6 THEN %><%=arrList(15,intLoop)%>
    		<%ELSEIF arrList(8,intLoop) = 8 THEN%><%=arrList(16,intLoop)%>
    		<%END IF%></a>
    	</td>
    	<td> 
    			<%IF chkState = 1 THEN%><input type="button" value="�ǽð�����" class="button" onClick="jsSetRealSale(<%=arrList(0,intLoop)%>,<%=chkState%>);"> <%END IF%>
    	</td>    			
    	<td>
    			<input type="button" value="��ǰ(<%=arrList(13,intLoop)%>)" class="button" onClick="javascript:jsGoURL('i',<%=arrList(0,intLoop)%>)">    		
    		</td>
    	<td><a href="javascript:jsMod(<%=arrList(0,intLoop)%>)" title="���� ��������"><%=FormatDate(arrList(10,intLoop),"0000.00.00")%></a></td>
    </tr>
    	
	<% Next
	ELSE
	%>
	<tr>
		<td colspan="13" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
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