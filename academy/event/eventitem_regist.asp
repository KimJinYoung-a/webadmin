<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �̺�Ʈ ��� - ��ǰ���
' History : 2010.09.29 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
Dim cEvtItem,cEvtCont,cEGroup ,eCode ,arrGroup ,iTotCnt, arrList,intLoop ,iPageSize, iCurrpage ,iDelCnt
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,estatedesc, ekinddesc ,strparm
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt ,strG, strSort	,sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind
	eCode = RequestCheckvar(Request("eC"),10)
	strG  = RequestCheckvar(Request("selG"),10)
	strSort  = RequestCheckvar(Request("selSort"),10)
	
	IF eCode = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
%>
		<script language="javascript">
			alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
			history.back();
		</script>
<%	
		dbget.close()	:	response.End
	END IF	
	
	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = RequestCheckvar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" Then
			iCurrpage = 1	
	END IF	  
		
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	'## �˻� #############################			
	sDate = RequestCheckvar(Request("selDate"),32)  '�Ⱓ 
	sSdate = RequestCheckvar(Request("iSD"),32)
	sEdate = RequestCheckvar(Request("iED"),32)
	
	sEvt = RequestCheckvar(Request("selEvt"),10)  '�̺�Ʈ �ڵ�/�� �˻�
	strTxt = Request("sEtxt")
	
	sCategory	= RequestCheckvar(Request("selC"),10) 'ī�װ�
	sState	 = RequestCheckvar(Request("eventstate"),10)'�̺�Ʈ ����	
	sKind = RequestCheckvar(Request("eventkind"),10)	'�̺�Ʈ����
		
	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################
	
'--�̺�Ʈ ����
set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'�̺�Ʈ �ڵ�
	
	cEvtCont.fnGetEventCont	 '�̺�Ʈ ���� ��������
	ekind 		=	cEvtCont.FEKind 
	ekinddesc	=	cEvtCont.FEKindDesc
	eman 		=	cEvtCont.FEManager 
	escope 		=	cEvtCont.FEScope 
	ename 		=	db2html(cEvtCont.FEName)
	esday 		=	cEvtCont.FESDay
	eeday 		=	cEvtCont.FEEDay
	epday 		=	cEvtCont.FEPDay
	elevel 		=	cEvtCont.FELevel
	estate 		=	cEvtCont.FEState
	estatedesc 	= cEvtCont.FEStateDesc
	eregdate 	=	cEvtCont.FERegdate 
set cEvtCont = nothing

'--�̺�Ʈ ��ǰ	
 set cEGroup = new ClsEventGroup
	cEGroup.FECode = eCode  	
	arrGroup = cEGroup.fnGetEventItemGroup		
set cEGroup = nothing

set cEvtItem = new ClsEvent	
	cEvtItem.FCPage = iCurrpage
	cEvtItem.FPSize = iPageSize	
	cEvtItem.FECode = eCode	
	cEvtItem.FESGroup = strG	
	cEvtItem.FESSort = strSort	
			
	arrList = cEvtItem.fnGetEventItem 		
	iTotCnt = cEvtItem.FTotCnt	'��ü ������  ��

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ �� 
%>

<script language="javascript">

// ������ �̵�
function jsGoPage(iP){
		document.fitem.iC.value = iP;		
		document.fitem.submit();	
}
	
// ����ǰ �߰� �˾�
function addnewItem(){
		var popwin;
		popwin = window.open("/academy/event/common/pop_event_additemlist.asp?eC=<%=eCode%>", "popup_item", "width=800,height=500,scrollbars=yes,resizable=yes");
		popwin.focus();
}
		
	
//����
function jsChSort(){
		document.fitem.submit();	
}

//�׷�˻�
function jsSearchGroup(){
		document.fitem.submit();	
}
	
//�׷��̵�	
function addGroup(){
		var frm,sValue,sGroup;		
		
		frm = document.fitem;
		sValue = "";
		sGroup =frm.eG.options[frm.eG.selectedIndex].value ;
				
		if(!frm.chkI) return;
		if(!sGroup){
		 alert("�̵��� �׷��� �����ϴ�.");
		 return;
		}
		
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
			sValue = frm.chkI.value;
		}
		
		if (sValue == "") {
			alert('���� ��ǰ�� �����ϴ�.');
			return;
		}
		
		document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
		document.frmG.itemidarr.value = sValue;
		document.frmG.submit();
}
	
	
	
//��ü����
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
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}

// ��ü �̹���ũ�� �ϰ� ��ȯ
function jsSizeChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.itemimgsize.length;i++){
			frm.itemimgsize[i].value=selv;
		}
	} else {
		frm.itemimgsize.value=selv;
	}
}
	
//����
function jsDel(sType, iValue){	
		var frm;		
		var sValue;		
		frm = document.fitem;
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

// ��ǰ ����/�̹��� ������ �ϰ� ����
function jsSortImgSize() {
	var frm;
	var sValue, sSort, sImgSize;
	frm = document.fitem;
	sValue = "";
	sSort = "";
	sImgSize = "";
		
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(!IsDigit(frm.sSort[i].value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sSort[i].focus();
				return;
			}
			
			if (sValue==""){
				sValue = frm.chkI[i].value;		
			}else{
				sValue =sValue+","+frm.chkI[i].value;		
			}	
			
			// ���ļ���
			if (sSort==""){
				sSort = frm.sSort[i].value;		
			}else{
				sSort =sSort+","+frm.sSort[i].value;		
			}

			// �̹��� ������
			if (sImgSize==""){
				sImgSize = frm.itemimgsize[i].value;		
			}else{
				sImgSize =sImgSize+","+frm.itemimgsize[i].value;		
			}	
		}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("���������� ���ڸ� �����մϴ�.");
			frm.sSort.focus();
			return;
		}
		sSort =  frm.sSort.value; 
		sImgSize =  frm.itemimgsize.value; 
	}

		document.frmSortImgSize.itemidarr.value = sValue;
		document.frmSortImgSize.sortarr.value = sSort;
		document.frmSortImgSize.sizearr.value = sImgSize;
		document.frmSortImgSize.submit();
}

	//�׷��߰�
	function jsAddGroup(){
		var winIG;
		winIG = window.open('iframe_eventitem_group.asp?ec=<%=eCode%>&T=1','popIG','width=700,height=600,scrollbars=yes');
		winIG.focus();
	}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10"> 
		<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
				<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
			</tr>	
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ekinddesc%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
			</tr>
			<tr>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
				<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
				<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
			</tr>			
		</table>
	</td>
</tr>	
		
<tr>
	<td >
		<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a">
			<form name="fitem" method="post" action="eventitem_regist.asp">
			<input type="hidden" name="iC" value="">
			<input type="hidden" name="eC" value="<%=eCode%>">
			<input type="hidden" name="menupos" value="<%=menupos%>">
			<input type="hidden" name="mode" value="">
			<input type="hidden" name="selGroup" value="">
			<tr align="center"  >
				<td align="left">  		        
		        	 �׷�˻�
		        	<select name="selG" onChange="jsSearchGroup();">
		        	<option value="">��ü</option>			        	
		       	<%IF isArray(arrGroup) THEN %>
		       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>������</option>
		       	<%	
		       		For intLoop = 0 To UBound(arrGroup,2)
		       	%>
		       		<option value="<%=arrGroup(0,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(0,intLoop)) THEN %> selected<%END IF%>> <%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)</option>
		    	<%	Next 
		    	END IF%>	
		       	</select> 			           			        	
		        </td>
		        <td align="right">
		         ���� : <select name="selSort" onchange="jsChSort();">			         			         		       					       		
		       		<option value="sitemid" >�Ż�ǰ��</option>			       					       		
		       		<option value="sevtitem" <%IF Cstr(strSort) = "sevtitem" THEN %>selected<%END IF%>>������</option>
		       		<option value="sbest" <%IF Cstr(strSort) = "sbest" THEN %>selected<%END IF%>>����Ʈ������</option>	
		       		<option value="shsell" <%IF Cstr(strSort) = "shsell" THEN %>selected<%END IF%>>�������ݼ�</option>			       	
		       		<option value="slsell" <%IF Cstr(strSort) = "slsell" THEN %>selected<%END IF%>>�������ݼ�</option>	
		       		<option value="sevtgroup" <%IF Cstr(strSort) = "sevtgroup" THEN %>selected<%END IF%>>�׷��</option>
		       		<option value="sbrand" <%IF Cstr(strSort) = "sbrand" THEN %>selected<%END IF%>>�귣��</option>
		       		</select>			       		
		        </td>			       
			</tr>
		</table>
	</td>
</tR>
		 
<tr>
	<td style="border-top:1px solid <%= adminColor("tablebg") %>;">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">												
		    <tr height="35">			      
		        <td align="left">       	
		       	<input type="button" value="���û���" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;      
		       	<select name="eG">
		       	<%IF isArray(arrGroup) THEN
		       		For intLoop = 0 To UBound(arrGroup,2)
		       	%>
		       		<option value=" <%=arrGroup(0,intLoop)%>" ><%IF arrGroup(5,intLoop) <> 0 THEN%>��&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)</option>
		    	<%	Next 
		    	  ELSE	
		    	%>
		    	<option value=""> --�׷����--</option>
		    	<%END IF%>	
		       	</select>
		       		<input type="button" value="���ñ׷��̵�" onClick="addGroup();" class="button">
		       			&nbsp; 	<input type="button" value="�׷��߰�" onClick="jsAddGroup();" class="button">   
		    	</td>
		    	<td align="right">
		    	<input type="button" value="����/������ ����" onClick="jsSortImgSize();" class="button">&nbsp; 
		       	<input type="button" value="����ǰ �߰�" onclick="addnewItem();" class="button">
		       	
		        </td>			      
		    </tr>
		</table>
	</td>
</tr>
		 
<tr>
	<td> 
		<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr bgcolor="#FFFFFF">
	   		<td colspan="15" align="left">�˻���� : <b><%=iTotCnt%></b>&nbsp;&nbsp;������ : <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
	   	</tr>
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
	    	<td>�׷��ڵ�</td>
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
	    	<td>����</td>
	    	<td>�̹���ũ��<br>
				<select name="selimgsize" onchange=jsSizeChg(this.value)>	
				<option value="1">100px</option>
				<option value="2">200px</option>
				</select>
	    	</td>
	    	<td>���</td>
	    </tr>
	    <%IF isArray(arrList) THEN 
	    	For intLoop = 0 To UBound(arrList,2)
	    %>
	    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';> 
	    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
	    	<td><%IF arrList(1,intLoop) <> 0 THEN%><%=arrList(1,intLoop)%><%END IF%></td>		    				    	
	    	<td>		    				    		
	    		<A href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
	    		<% if cEvtItem.IsSoldOut(arrList(14,intLoop),arrList(16,intLoop),arrList(20,intLoop),arrList(21,intLoop)) then %>
	    			<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
	    		<% end if %>
	    	</td>
	    	<td>
	    		<% if (Not IsNull(arrList(12,intLoop)) ) and (arrList(12,intLoop)<>"") then %>
					<img src="<%= imgFingers %>/diyItem/webimage/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(12,intLoop)%>">
				<%end if%>
	    	</td>    	
	    	<td><%=db2html(arrList(3,intLoop))%></td>
	    	<td align="left"><%=db2html(arrList(4,intLoop))%></td>
	    	<td>
	    		<%
				Response.Write FormatNumber(arrList(7,intLoop),0)
				'���ΰ�
				if arrList(18,intLoop)="Y" then
					Response.Write "<br><font color=#F08050>(��)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
				end if
				'������
				if arrList(22,intLoop)="Y" then
					Select Case arrList(23,intLoop)
						Case "1"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(7,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
						Case "2"
							Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(7,intLoop)-arrList(24,intLoop),0) & "</font>"
					end Select
				end if
				%>
			</td>
	    	<td>
	    		<%
				Response.Write FormatNumber(arrList(8,intLoop),0)
				'���ΰ�
				if arrList(18,intLoop)="Y" then
					Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(10,intLoop),0) & "</font>"
				end if
				'������
				if arrList(22,intLoop)="Y" then
					if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
						if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
							Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
						else
							Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(25,intLoop),0) & "</font>"
						end if
					end if
				end if
			%>
			</td>
	    	<td><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>    	
	    	<td><%= fnColor(arrList(14,intLoop),"yn") %></td>
	    	<td><%= fnColor(arrList(19,intLoop),"yn") %></td>
	    	<td><%= fnColor(arrList(16,intLoop),"yn") %></td>    				    	
	    	<td><input type="text" name="sSort" value="<%=arrList(2,intLoop)%>" size="4" style="text-align:right;"></td>
	    	<td><% sbGetOptEventCodeValue "itemimgsize", arrList(27,intLoop), false, "" %></td>
	    	<td><input type="button" value="����" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td>	    	
	    </tr>   
	   <%	Next
	   	ELSE
	   %>
	   	<tr  align="center" bgcolor="#FFFFFF">
	   		<td colspan="15">��ϵ� ������ �����ϴ�.</td>
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
		        <td width=50 align="right"><a href="event_list.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
		    </tr>			  
		</form>    
		</table>
	</td>
</tr>
</table>

<!-- �׷��̵�--->
<form name="frmG" method="post" action="eventitem_process.asp">
	<input type="hidden" name="mode" value="G">
	<input type="hidden" name="iC" value="<%=iCurrpage%>">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="selG" value="<%=strG%>">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="selGroup" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- ���û���--->
<form name="frmDel" method="post" action="eventitem_process.asp">
	<input type="hidden" name="mode" value="D">
	<input type="hidden" name="iC" value="<%=iCurrpage%>">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="selG" value="<%=strG%>">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- ���� �� �̹���ũ�� ����--->
<form name="frmSortImgSize" method="post" action="eventitem_process.asp">
	<input type="hidden" name="mode" value="S">
	<input type="hidden" name="eC" value="<%=eCode%>">
	<input type="hidden" name="selG" value="<%=strG%>">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="sortarr" value="">
	<input type="hidden" name="sizearr" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>

<%
	set cEvtItem = nothing
%>	

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
