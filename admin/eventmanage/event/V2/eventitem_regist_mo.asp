<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventitem_regist_mo.asp
' Description :  �̺�Ʈ ��� - ����� ��ǰ���
' History : 2007.02.21 ������ ����
'           2008.10.20 ��ǰ�̹��� ũ�� �߰�(������)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<%
'��������
Dim eCode
Dim cEvtItem,cEvtCont,cEGroup
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,estatedesc, ekinddesc
Dim arrGroup,arrGroup_mo
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
Dim iDispYCnt, iDispNCnt

Dim strG, strSort
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind,eisort
Dim strparm, Brand
dim makerid, itemname, itemid
dim  itemsort,eItemListType, blnWeb, blnMobile, blnApp
dim eChannel
 
  
strG  		= requestCheckvar(Request("selG"),10)
strSort  	= requestCheckvar(Request("selSort"),1)
	
eCode 		= requestCheckvar(request("eC"),10)
itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
itemsort  	= requestCheckvar(request("itemsort"),32)
eChannel    = requestCheckvar(request("eCh"),1)

	IF eCode = "" THEN	'�̺�Ʈ �ڵ尪�� ���� ��� back
%>
		<script language="javascript">
		<!--
			alert("���ް��� ������ �߻��Ͽ����ϴ�. �����ڿ��� �������ֽʽÿ�");
			history.back();
		//-->
		</script>
	<%	dbget.close()	:	response.End
	END IF	
	if eChannel = "" then eChannel = "M"
	if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	itemid = left(arrItemid,len(arrItemid)-1)
end if

	'�Ķ���Ͱ� �ޱ� & �⺻ ���� �� ����
	iCurrpage = Request("iC")	'���� ������ ��ȣ
     
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	
 
	iPageSize = 30		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	
	'## �˻� #############################			
	sDate = Request("selDate")  '�Ⱓ 
	sSdate = Request("iSD")
	sEdate = Request("iED")	
	
	sEvt = Request("selEvt")  '�̺�Ʈ �ڵ�/�� �˻�
	strTxt = Request("sEtxt")
	
	sCategory	= Request("selC") 'ī�װ�
	sState	 = Request("eventstate")'�̺�Ʈ ����	
	sKind = Request("eventkind")	'�̺�Ʈ����
	 
	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################
	
	'������ ��������
	
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
		estatedesc 	= 	cEvtCont.FEStateDesc
		eregdate 	=	cEvtCont.FERegdate
		blnWeb      =   cEvtCont.FIsWeb
        blnMobile   =   cEvtCont.FIsMobile
        blnApp      =   cEvtCont.FIsApp
        
		'�̺�Ʈ ȭ�鼳�� ���� ��������
		cEvtCont.fnGetEventDisplay		
		Brand		= 	cEvtCont.FEBrand	
		eisort 		=   cEvtCont.FEISort	
		eItemListType =   cEvtCont.FEListType
	set cEvtCont = nothing
	
	'--�̺�Ʈ ��ǰ	
	 set cEGroup = new ClsEventGroup
 		cEGroup.FECode = eCode  
 		cEGroup.FEChannel = eChannel
  		arrGroup = cEGroup.fnGetEventItemGroup	 
 	set cEGroup = nothing
 	
 	if itemsort = "" then itemsort = eisort
	set cEvtItem = new ClsEvent	
		
		cEvtItem.FPSize = iPageSize	
		cEvtItem.FECode = eCode	
		cEvtItem.FRectMakerid = makerid
		cEvtItem.FRectItemid = itemid
		cEvtItem.FRectItemName = itemname  
       
        cEvtItem.FCPage = iCurrpage
        cEvtItem.FESGroup = strG	
		cEvtItem.FESSort = itemsort	
        cEvtItem.FEChannel = eChannel
 		arrList = cEvtItem.fnGetEventItem 		
 		iTotCnt = cEvtItem.FTotCnt	'��ü ������  ��
        iDispYCnt = cEvtItem.FDispYMCnt
        iDispNCnt = cEvtItem.FDispNMCnt 
        
 	    iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ �� 
 	     
%>
<script language="javascript">
<!--
// ������ �̵�
function jsGoPage(iCurrpage){
    document.fitem.iC.value = iCurrpage;		
    document.fitem.submit();	
}
	
// ����ǰ �߰� �˾�
function addnewItem(eChannel){
		var popwin;
		popwin = window.open("pop_event_additemlist.asp?eC=<%=eCode%>&makerid=<%= Brand %>&egcode=<%=strG%>&eCh="+eChannel, "popup_item", "width=1024,height=768,scrollbars=yes,resizable=yes");
		popwin.focus();
}
		
	
//����
function jsChSort(){ 
		document.fitem.submit();	
}

//�׷�˻�
function jsSearch(){ 
		document.fsearch.submit();	
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
		
		if(confirm("�׷��̵��� PC-Web �� Mobile/App �� �Բ� �̵��˴ϴ�. �̵��Ͻðڽ��ϱ�?")){
		document.frmG.selGroup.value = frm.eG.options[frm.eG.selectedIndex].value;
		document.frmG.itemidarr.value = sValue;
		document.frmG.submit();
	   }
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

// ��ü ���ÿ���  �ϰ� ��ȯ
function jsDispChg(selv) { 
    if(selv=="") {return;}
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		 for (var i=0;i<frm.chkI.length;i++){
		    if (selv=="Y"){ 
			  eval("frm.eDisp_"+i)[0].checked = true;
			  eval("frm.eDisp_"+i)[1].checked = false;
		   }else{
		      eval("frm.eDisp_"+i)[0].checked = false;
		      eval("frm.eDisp_"+i)[1].checked = true;
		   }
		}
	} else {
	    if(selv=="Y") {
		  frm.eDisp_0[0].checked=true;
		  frm.eDisp_0[1].checked=false;
		}else{
		  frm.eDisp_0[0].checked=false;
		  frm.eDisp_0[1].checked=true;
		}
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
	var sValue, sSort, sImgSize,sUsing, sSort_mo, sImgSize_mo,sUsing_mo,sDisp;
	frm = document.fitem;
	sValue = "";
	sSort = ""; 
	sDisp = ""
	var itemid;	
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){ 
			if (frm.chkI[i].checked){
			if(!IsDigit(frm.sSort[i].value)){
				alert("���������� ���ڸ� �����մϴ�.");
				frm.sSort[i].focus();
				return;
			}
			 
			
		  itemid = frm.chkI[i].value;		
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
   
			//���ÿ���
 			if(sDisp ==""){
 			   if (eval("frm.eDisp_"+i)[0].checked ==true){
 			      sDisp =   eval("frm.eDisp_"+i)[0].value;
 			    }else{
 			       sDisp =   eval("frm.eDisp_"+i)[1].value;
 			    }
 			}else{
 			    if (eval("frm.eDisp_"+i)[0].checked ==true){
 			      sDisp =  sDisp+","+ eval("frm.eDisp_"+i)[0].value;
 			    }else{
 			       sDisp =  sDisp+","+ eval("frm.eDisp_"+i)[1].value;
 			    } 
 		    } 
		}
	}
	}else{
		sValue = frm.chkI.value;
		if(!IsDigit(frm.sSort.value)){
			alert("���������� ���ڸ� �����մϴ�.");
			frm.sSort.focus();
			return;
		} 
		 
		sSort   = frm.sSort.value ;
       
        if (frm.eDisp_0[0].checked ==true){
 		    sDisp =  frm.eDisp_0[0].value;
 		}else{
 		    sDisp =  frm.eDisp_0[1].value;
        }
	}

		document.frmSortImgSize.itemidarr.value = sValue;
		document.frmSortImgSize.sortarr.value = sSort; 
		document.frmSortImgSize.disparr.value = sDisp;
	 	document.frmSortImgSize.submit();
}

	//�׷��߰�
	function jsAddGroup(eCode,gCode,eChannel){
		var winIG; 
		winIG = window.open('pop_eventitem_group.asp?eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel+'&sTarget=item','popG','width=600, height=500,scrollbars=yes,resizable=yes'); 
		winIG.focus();
	}
	 
	 
	//�����/�� ��ǰ����Ʈ��Ÿ�� ����
	function jsChangeListType(){
	    var i, eilt; 
	    for(i=0;i<document.fitem.itemlisttype.length;i++){ 
	        if(document.fitem.itemlisttype[i].checked){
	            eilt = document.fitem.itemlisttype[i].value;
	        }
	    }
 	  
	    document.frmLT.eILT.value = eilt;
	    document.frmLT.submit();
	}
	
	 function jsChkTrue(i){ 
	   if ( document.fitem.chkI.length > 1){ 
	     document.fitem.chkI[i].checked = true;
	}else{
	     document.fitem.chkI.checked =true;
	}
	}
//-->
</script>
<style type="text/css">
div.btmLine {background:url(/images/partner/admin_grade.png) left bottom repeat-x; padding-bottom:5px;}    
.tab {position:relative; z-index:50;}
.tab ul {_zoom:1; border-left:1px solid #ccc; border-bottom:1px solid #ccc; list-style:none; margin:0; padding:0;}
.tab ul:after {content:""; display:block; height:0; clear:both; visibility:hidden;}
.tab ul li {float:left; text-align:center; height:23px;padding-top:7px;border:1px solid #ccc; margin:0 0 -1px -1px; cursor:pointer;   background-color:#fff; }
.tab ul li.selected {background-color:#e3f1fb; position:relative; font-weight:bold;}
.col11 {width:15% !important;}
</style>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  style="padding-top:10px">
    <tr>
		<td style="padding-bottom:10"> 
			<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ�ڵ�</td>
					<td width="30%" bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=eCode%>&nbsp;&nbsp; <a href="<%=vmobileUrl%>/event/eventmain.asp?eventid=<%=eCode%>" target="_blank">[�̸�����]</a></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=ename%></td>
				</tr>	
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
						<%=ekinddesc%>
						<% if ekind="16" then %>
							(<%= brand %>)
						<% end if %>
					</td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=estatedesc%></td>
				</tr>
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=esday%>~ <%=eeday%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">��÷ ��ǥ��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=epday%></td>
				</tr>		
				<tr>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��ǰ ����</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=fnGetEventCodeDesc("itemsort", eisort)%></td>
					<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">ä��</td>
					<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%if blnWeb then%>PC-WEB <%END IF%><%if blnMobile then%>Mobile <%END IF%><%if blnApp then%>App <%END IF%></td>
				</tr>
			</table>
		</td>
	</tr>	
	<tr>
	    <td>
	        <div class="tab btmLine">
				<ul style="margin-left:-1px">
					<li class="col11"  onclick="location.href='eventitem_regist.asp?eC=<%=eCode%>&menupos=<%=menupos%>'">PC_WEB ��ǰ���</li>
            		<li class="col11 selected">Moblie/ APP ��ǰ���</li>
				</ul>
			</div>
		</td>
	</tr>	
	<tr><!-- �˻�--->
		<td>
		    <table cellspacing="5"  bgcolor="#e3f1fb" width="100%" class="a" cellpadding="0">
		        <tr>
		            <td bgcolor="#FFFFFF">	 
            			<form name="fsearch" method="post" action="eventitem_regist_mo.asp"> 
            				<input type="hidden" name="eC" value="<%=eCode%>">
            				<input type="hidden" name="eCh" value="<%=eChannel%>">
            				<input type="hidden" name="menupos" value="<%=menupos%>">
            				<input type="hidden" name="mode" value="">
            				<input type="hidden" name="selGroup" value="">
            				<input type="hidden" name="itemsort" value="<%=itemsort%>">
            			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
            				<tr align="center" >
            					<td  width="100" bgcolor="<%= adminColor("tabletop") %>">�˻� ����</td>
            					<td align="left"  bgcolor="#ffffff">  	 
            						<table border="0" cellpadding="1" cellspacing="1" class="a">
                						<tr>
                							<td style="white-space:nowrap;">�׷�: 
                								<select name="selG" onChange="jsSearch();" class="select">
                						        	<option value="">��ü</option>			        	
                						       	<%IF isArray(arrGroup) THEN %>
                						       		<option value="0"  <%IF Cstr(strG) = "0" THEN%>selected<%END IF%>>������</option>
                						       	<%	dim sumi, i
                						       		For intLoop = 0 To UBound(arrGroup,2)
                						       		 sumi = 0
                						       	%>
                						       		<option value="<%=arrGroup(9,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(9,intLoop)) THEN %> selected<%END IF%>  <%if not arrGroup(8,intLoop) then%>style="color:gray;"<%end if%>>
                						       		    <%IF arrGroup(5,intLoop) <> 0 THEN%>��&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)
                						       		    <% if intLoop < UBound(arrGroup,2)  then 
					                                       for i = 1 to (UBound(arrGroup,2)-intLoop) 
                                    					     if arrGroup(9,intLoop) = arrGroup(9,intLoop+i) then
                                    					        sumi = sumi + 1  
                                    					         %>
                                    					    + <%=arrGroup(0,intLoop+i)%>(<%=arrGroup(1,intLoop+i)%>)
                                    					<%   else 
                                    					        exit for
                                    					     end if 
                                    					    next
                                    					   end if 
                                    					     %>
                						       		    <%if not arrGroup(8,intLoop) then%> -[���þ���]<%end if%>
                						       		 </option>
                						    	<%	 intLoop = intLoop+sumi
                						    	    Next 
                						    	END IF%>	
                						       	</select> 	
                			       			</td> 
                							<td style="white-space:nowrap;padding-left:10px;">�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>  
                							<td style="white-space:nowrap;padding-left:10px;">��ǰ�ڵ�:</td>
                							<td style="white-space:nowrap;" rowspan="2"><textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea> </td>
                						</tr> 
                						<tr>
                						    <td colspan="4">��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
            						</tr>
            			        	</table>        			        	
            			        </td>
            			        <td   width="100" bgcolor="<%= adminColor("gray") %>">
            						<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
            					</td>
            			    </tr> 
            			</table>
            			</form>
            		</td>
            	</tR> <!-- �˻�---> 
            	<tr>
            		<td style="padding-top:10px;" valign="top">  
            		     <div id="divMA">
            		       <form name="fitem" method="post" action="eventitem_regist_mo.asp">
                           <input type="hidden" name="menupos" value="<%=menupos%>">
                           <input type="hidden" name="mode" value="">
                           <input type="hidden" name="iC" value="">
                           <input type="hidden" name="eC" value="<%=eCode%>"> 
                           <input type="hidden" name="eCh" value="<%=eChannel%>"> 
                           <input type="hidden" name="selGroup" value="">
                           <input type="hidden" name="selG" value="<%=strG%>">
                           <input type="hidden" name="makerid" value="<%=makerid%>">
                           <input type="hidden" name="itemname" value="<%=itemname%>">
                           <input type="hidden" name="itemid" value="<%=itemid%>">
 
            		        <table width="100%" border="0" align="center" cellpadding="0"  class="a" cellspacing="1">	  
            		            <tr>
            		                <td>
                                 	    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0" class="a" >		
                                 	        <tr>
                                 	           <td colspan="2">��ǰ ����Ʈ ��Ÿ��:	
                                 		     		<input type="radio" name="itemlisttype"  value="1"  <%IF eItemListType = "1" THEN%>checked<%END IF%>>������ 
                                 				    <input type="radio" name="itemlisttype"  value="2"  <%IF eItemListType = "2" THEN%>checked<%END IF%>>����Ʈ�� 
                                 				    <input type="radio" name="itemlisttype"  value="3"  <%IF eItemListType = "3" THEN%>checked<%END IF%>>BIG��	
                                 				    &nbsp;&nbsp;<input type="button" value="����" class="button" onClick="jsChangeListType();" style="width:60px;">
                                 			    </td> 
                                 	        </tr>
                                 		    <tr>
                                 		        
                                     		      <td>
                                 	                <input type="button" value="���û���" onClick="jsDel(0,'');" class="button">&nbsp;&nbsp;&nbsp;    
                                     		     	<select name="eG" class="select">
                                     		     	<%IF isArray(arrGroup) THEN %>
                          			            	<option value=""> --����--</option>
                          			       	    <%
                						       		    For intLoop = 0 To UBound(arrGroup,2)
                						       		     
                						       	%>
                						       		<option value="<%=arrGroup(0,intLoop)%>" <%IF Cstr(strG) = Cstr(arrGroup(0,intLoop)) THEN %> selected<%END IF%>  <%if not arrGroup(8,intLoop) then%>style="color:gray;"<%end if%>>
                						       		    <%IF arrGroup(5,intLoop) <> 0 THEN%>��&nbsp;<%END IF%><%=arrGroup(0,intLoop)%>(<%=arrGroup(1,intLoop)%>)
                						       		     
                						       		    <%if not arrGroup(8,intLoop) then%> -[���þ���]<%end if%>
                						       		 </option>
                						    	<%	  
                						    	    Next 
                						     
                                         		  	  ELSE	
                                         		  	%>
                                     		  	        <option value=""> --�׷����--</option>
                                     		  	    <%END IF%>	
                                 		     	    </select>   
                                 		     		<input type="button" value="���ñ׷��̵�" onClick="addGroup();" class="button">
                                 		     			&nbsp; 	<input type="button" value="�׷��߰�" onClick="jsAddGroup('<%=eCode%>','','<%=eChannel%>');" class="button">   
                                 		     		 
                                 		  	    </td>	
                                 		  	    <td align="right"> 
                                 		            <input type="button" value="���û�ǰ ����" onClick="jsSortImgSize();" class="button">&nbsp; 
                                 		     	    <input type="button" value="����ǰ �߰�" onclick="addnewItem('<%=eChannel%>');" class="button"> 
                                                </td>			      
                                 		    </tr>
                                 		</table>
                                    </td>
                                </tr> 
                                <tr>
                      	            <td> 
                      			        <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
                                         <tr bgcolor="#FFFFFF">
                                     	    <td  colspan="20" >
                                     	        <table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="0" > 
                                     	            <tr>
                                         		        <td align="left">[�˻����] <font color="blue">���� Y: <%=iDispYCnt%></font> /  <font color="red">���� N: <%=iDispNCnt%></font> / <b>��: <%=iTotCnt%></b>&nbsp;&nbsp;&nbsp;&nbsp;������: <b><%=iCurrpage%> / <%=iTotalPage%></b></td>
                                         		        <td align="right">���� : <%sbGetOptCommonCodeArr  "itemsort",itemsort,False,False,"onchange='jsChSort();'"%></td>
                                         		    </tr>
                                         		</table>
                                         	</td>
                                     	</tr>
                                        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                                     		<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
                                     		<td>�׷��ڵ�</td>
                                     		<td>��ǰID</td>
                                     		<td>�̹���</td>
                                     		<td>�귣��</td>
                                     		<td>��ǰ��</td>
                                     		<td>�ǸŰ�</td>
                                     		<td>���԰�</td>
                                     		<td>������</td>
                                     		<td>���</td>	
                                     		<td>�Ǹſ���</td>	
                                     		<td>��ǰ��뿩��</td>	
                                     		<td>��������</td> 
                                     		<td>����</td>  
                                     		<td> <select name="selDisp" class="select" onChange="jsDispChg(this.value);">
                      			    	        <option value="">���ÿ���</option>
                      			    	        <option value="Y">Y</option>
                      			    	        <option value="N">N</option>
                      			    	    </select></td> 
                                     	</tr> 
                                     		<%IF isArray(arrList) THEN 
                                     			For intLoop = 0 To UBound(arrList,2)
                                     		%>
                                     	<tr align="center" bgcolor="<%if  arrList(29,intLoop) then%>#FFFFFF<%else%>gray<%end if%>">    
                                     		<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
                                     		<td><%IF arrList(1,intLoop) <> 0 THEN%><%=arrList(1,intLoop)%><%END IF%></td>		    				    	
                                     		<td>
                                     				<!-- 2007/05/05 ������ ���� -- ǰ�� ǥ�� -->			    		
                                     				<A href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a>
                                     				<% if cEvtItem.IsSoldOut(arrList(14,intLoop),arrList(16,intLoop),arrList(20,intLoop),arrList(21,intLoop)) then %>
                                     					<br><img src="http://webadmin.10x10.co.kr/images/soldout_s.gif" width="30" height="12">
                                     				<% end if %>
                                     		</td>
                                     		<td><% if (Not IsNull(arrList(12,intLoop)) ) and (arrList(12,intLoop)<>"") then %>
                                     		     <img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(12,intLoop)%>">
                                     		    <%end if%>
                                     		</td>    	
                                     		<td><%=db2html(arrList(3,intLoop))%></td>
                                     		<td align="left">&nbsp;<%=db2html(arrList(4,intLoop))%></td>
                                     		<td><%
                                     			Response.Write FormatNumber(arrList(7,intLoop),0)
                                     			'���ΰ�
                                     			if arrList(18,intLoop)="Y" then
                                     				Response.Write "<br><font color=#F08050>(��)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
                                     			end if
                                     			'������
                                     			if arrList(22,intLoop)="Y" then
                                     				Select Case arrList(23,intLoop)
                                     					Case "1"
                                      						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
                                     					Case "2"
                                     						Response.Write "<br><font color=#5080F0>(��)" & FormatNumber(arrList(5,intLoop)-arrList(24,intLoop),0) & "</font>"
                                     				end Select
                                     			end if
                                     		%></td>
                                     		<td><%
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
                                     		%></td>
                                     			<td><%if arrList(18,intLoop)="Y" then%>
                                     						<font color=#F08050><%=formatnumber(((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop))*100,0)%>%</font>
                                     						 
                                     						<%end if%>
                                     				<%if arrList(22,intLoop)="Y" then 
                      						if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
                      					        if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
                      						         Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(8,intLoop),0) & "</font>"
                      					        else
                      						        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(24,intLoop),0) 
                      						         if arrList(23,intLoop)="1" then 
                      						         Response.Write "%"
                      						        else
                      						         Response.Write "��"
                      						        end if
                      						         Response.Write "</font>"
                      					        end if
                      				        end if
                      						 end if%>		
                                     			</td>
                                     		   	<td><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>    	
                                     			<td><%= fnColor(arrList(14,intLoop),"yn") %></td>
                                     			<td><%= fnColor(arrList(19,intLoop),"yn") %></td>
                                     			<td><%= fnColor(arrList(16,intLoop),"yn") %></td>    				    	
                                     			<td><input type="text" name="sSort" value="<%=arrList(2,intLoop)%>" size="4" style="text-align:right;"></td> 
                                     			<td><input type="radio" name="eDisp_<%=intLoop%>" value="1" <%if arrList(29,intLoop) then%>checked<%end if%> onClick="jsChkTrue('<%=intLoop%>');"><%if arrList(29,intLoop) then%><font color="#5080F0"><%end if%>Y </font>
                      			    	    <input type="radio" name="eDisp_<%=intLoop%>" value="0" <%if not arrList(29,intLoop) then%>checked<%end if%> onClick="jsChkTrue('<%=intLoop%>');"><%if not arrList(29,intLoop) then%><font color="red"><%end if%>N</font></td>
                                     			<!--td><input type="button" value="����" onClick="jsDel(1,<%=arrList(0,intLoop)%>);" class="button"></td-->	
                                     		</tr>   
                                     			   <%	Next
                                     			   	ELSE
                                     			   %>
                                     		<tr  align="center" bgcolor="#FFFFFF">
                                     			<td colspan="19">��ϵ� ������ �����ϴ�.</td>
                                     		</tr>	
                                     			   <%END IF%>
                                     	</table>
                                    </td>
                               </tr>
                               <tr>
                                    <td> <!-- ����¡ó�� -->
                                          <%		
                                     	iStartPage = (Int((iCurrpage-1)/iPerCnt)*iPerCnt) + 1	
                                     	
                                     	If (iCurrpage mod iPerCnt) = 0 Then																
                                     		iEndPage = iCurrpage
                                     	Else								
                                     		iEndPage = iStartPage + (iPerCnt-1)
                                     	End If	
                                     	%>
                                     	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  >
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
                                    </td>
                                </tr>
                            </table> 
                        </div> 
            		</td>
            	</tR> 
	        </table>
	    </td>
	</tr>
	<tr>
	    <td   align="right"><a href="index.asp?menupos=<%=menupos%>&<%=strparm%>"><img src="/images/icon_list.gif" border="0"></a></td>			        
	 </tr>
</table>
<%
	set cEvtItem = nothing
%>	
<!-- �׷��̵�--->
<form name="frmG" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="G">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="selGroup" value="">
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- ���û���--->
<form name="frmDel" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="D">
<input type="hidden" name="iC" value="<%=iCurrpage%>">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="eCh" value="<%=eChannel%>">
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
<input type="hidden" name="disparr" value="">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="itemid" value="<%=itemid%>"> 
<input type="hidden" name="itemsort" value="<%=itemsort%>">   
<input type="hidden" name="eCh" value="<%=eChannel%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
 
<form name="frmLT" method="post" action="eventitem_process.asp">
<input type="hidden" name="mode" value="L">
<input type="hidden" name="eC" value="<%=eCode%>">
<input type="hidden" name="eILT" value="">
<input type="hidden" name="selG" value="<%=strG%>">
<input type="hidden" name="makerid" value="<%=makerid%>">
<input type="hidden" name="itemname" value="<%=itemname%>">
<input type="hidden" name="itemid" value="<%=itemid%>"> 
<input type="hidden" name="itemsort" value="<%=itemsort%>">   
<input type="hidden" name="eCh" value="<%=eChannel%>">   
<input type="hidden" name="menupos" value="<%=menupos%>">
</form>
<!-- ǥ �ϴܹ� ��-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
