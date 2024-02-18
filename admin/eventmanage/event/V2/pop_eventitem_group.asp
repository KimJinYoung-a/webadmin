<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/pop_eventitem_group.asp
' Description :  �̺�Ʈ �׷���
' History : 2007.02.22 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V2.asp"-->
<%
Dim eCode : eCode = Request("eC")
Dim eGCode : eGCode = Request("eGC")
Dim vYear : vYear = Request("yr") 
Dim sTarget : sTarget = request("sTarget")
dim eChannel : eChannel = requestCheckVar(Request("eCh"),1)
Dim cEGroup, arrP,intP,sM
Dim gpcode, gdesc, gsort, gimg,gdepth,gpdesc,glink
Dim arrImg, slen, sImgName
Dim arrList,intg  , gdisp
 gdisp = True
 set cEGroup = new ClsEventGroup
 	cEGroup.FECode = eCode
 	cEGroup.FEChannel = eChannel
 	cEGroup.FGDisp = 1
  	arrP = cEGroup.fnGetRootGroup
  	sM = "I"
  	IF (eGCode <> "" and eGCode <> "0" and not isnull(eGCode)) THEN
	  	cEGroup.FEGCode = eGCode
	  	cEGroup.fnGetEventItemGroupCont		
	  	gpcode 	= cEGroup.FGPCode
	  	gdesc  	= cEGroup.FGDesc
	  	gsort	= cEGroup.FGSort
	  	gdepth	= cEGroup.FGDepth
	  	gpdesc  = cEGroup.FGPDesc 
	  	gdisp   = cEGroup.FGDisp
	  	sM = "U"
	END IF  	 
 	 	 
  	arrList = cEGroup.fnGetEventItemGroup		
  	vYear = cEGroup.FRegdate
 set cEGroup = nothing
if gsort = "" then gsort = 0

%> 
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript">
<!--
 function jsGroupSubmit(){
 	if(!document.frmG.sGD.value){
 	alert("�׷���� �Է����ּ���");
 	return ;
 	}
 	document.frmG.submit();
 } 

 function jsSetGroup(eCode,gCode,eChannel, sTarget){
 	 self.location.href = 'pop_eventitem_group.asp?eC='+eCode+'&eGC='+gCode+'&eCh='+eChannel+'&sTarget='+sTarget ;
 } 
 
 function jsDelGroup(eCode,gCode,eChannel,sTarget){
 	if(confirm("������ �׷��� PC-WEb, Mobile/App�� ���ÿ� ���ǰ� �ִ� �׷��Դϴ�. ���� ó���� �� ä�ο��� ��� �����˴ϴ�. \n\n�����Ͻðڽ��ϱ�?")){
 		
 		document.frmG.mode.value = "D";
 		document.frmG.eC.value = eCode;
 		document.frmG.eGC.value = gCode;
 		document.frmG.eCh.value = eChannel;
 		document.frmG.sTarget.value = sTarget;
 		document.frmG.submit();	 
 	}
 } 
 
 
 
$(document).ready(function(){  
  $("#btn_g").click(function(){   
  	<%if sTarget = "item" THEN%>
  	opener.location.reload();
  	<%else%>
  	    <%if eChannel ="P" THEN%>
  	    $("#divFrm3", opener.document).html($("#divIpG").html()); 
  	    <%elseif eChannel ="M" THEN%>  
  	    $("#divMFrm3", opener.document).html($("#divIpMG").html());
  	    <%END IF%>
  	<%end if%>
  	window.close();
  }); 
  
  $( "#subList" ).sortable({
		placeholder: "ui-state-highlight",
		start: function(event, ui) {
			ui.placeholder.html('<td height="54" colspan="10" style="border:1px solid #F9BD01;">&nbsp;</td>');
		},
		stop: function(){
			var i=99999;
			$(this).parent().find("input[name^='eSArr']").each(function(){
				if(i>$(this).val()) i=$(this).val()
			});
			if(i<=0) i=1;
			$(this).parent().find("input[name^='eSArr']").each(function(){
				$(this).val(i);
				i++;
			}); 
		}
	});
});

	
	//�׷���� ��������
	function jsChSort(){
	    var sGDarr ="";
	    if(confirm("����� ������ �����Ͻðڽ��ϱ�? �׷쿡 ���Ե� ��ǰ�� ���� �̵��մϴ�.")){ 
	        if(typeof(document.frmL.eGCArr.length)=="undefined"){
	              sGDarr =document.frmL.eGD.value.replace(/[|]/g," ");  
	        }else{ 
    	        for (var i=0;i<document.frmL.eGCArr.length;i++){ 
    	          
    	            if(sGDarr == ""){
    	                 sGDarr =document.frmL.eGD[i].value.replace(/[|]/g," "); 
    	            }else{   
    	            sGDarr = sGDarr + "|" + document.frmL.eGD[i].value.replace(/[|]/g," ");
    	            } 
    	        }   
	        
	    }  
	       document.frmL.sGDarr.value = sGDarr;
	        document.frmL.submit();
	    }
	}
	
	//�׷���ġ��
	function jsSetJoin(){
	    var eGCArr =""; 
	    var icount = 0;
	    for (var i=0;i<document.frmL.chks.length;i++){
	        if(document.frmL.chks[i].checked){
	            if(eGCArr == ""){
	                 eGCArr =document.frmL.eGCMoArr[i].value; 
	            }else{   
	            eGCArr = eGCArr + "," + document.frmL.eGCMoArr[i].value;
	            }
	            icount=icount+1;
	        }
	    }
	    
       if (icount <= 1){
 	        alert("�׷��� 2�� �̻� �������ּ���");
 	       return;
 	    }
	    document.frmGM.eGCArr.value = eGCArr;
	    document.frmGM.mode.value = "J";
	    document.frmGM.submit();
	}
	
	//�׷쳪����
	function jsSetDivide(gCode){
	    if(confirm("�׷��� �����ðڽ��ϱ�?")){  
	    document.frmGM.eGC.value = gCode;
	    document.frmGM.mode.value = "Di";
	    document.frmGM.submit(); 
    	}
	}
	
	function jsDispGroup(gCode,isDisp){
     var strMsg ="";
      
    if (isDisp=="0"){   
      strMsg = "���ü����� [N-���þ���]���� �����Ͻðڽ��ϱ�?";
    }else{
      strMsg = "���ü����� [Y-������]���� �����Ͻðڽ��ϱ�?";
    }
    
    if(confirm(strMsg)){
 	    document.frmGM.eGC.value = gCode;
	    document.frmGM.mode.value = "A";
	    document.frmGM.eIsDisp.value = isDisp;
	    document.frmGM.submit();  
 	}
}
//-->
</script>
<form name="frmGM" method="post" action="eventgroup_process.asp">
    <input type="hidden" name="eC" value="<%=eCode%>"> 
	<input type="hidden" name="mode" value="">  
	<input type="hidden" name="eCh" value ="<%=eChannel%>"> 
	<input type="hidden" name="sTarget" value ="<%=sTarget%>">  
	<input type="hidden" name="eGCArr" value="">
	<input type="hidden" name="eGC" value="">
	<input type="hidden" name="eIsDisp" value="">
</form>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> �̺�Ʈ �׷� ���</div><hr>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
 <tr>
 	<td>
 		<form name="frmG" method="post" action="eventgroup_process.asp"   onSubmit="return jsGroupSubmit(this);">
		<input type="hidden" name="eC" value="<%=eCode%>">
		<input type="hidden" name="eGC" value="<%=eGCode%>">
		<input type="hidden" name="mode" value="<%=sM%>">  
		<input type="hidden" name="eCh" value ="<%=eChannel%>"> 
		<input type="hidden" name="sTarget" value ="<%=sTarget%>"> 
		<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0"> 
			<tr>
				<td> 
					<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
					    <tr height="30" > 
            				<%IF eChannel ="M" then%>
            				<td bgcolor="#e3f1fb" align="center" colspan="2"><b>Mobile / App</b></td>
            				<%ELSE%>
            				<td bgcolor="#FAECC5" align="center" colspan="2"><b>PC-WEB</b></td>
            				<%END IF%>
            			</tr>
						<tr>
							<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����׷�</td>
							<td bgcolor="#FFFFFF">
							<%IF gdepth = "" THEN%>
							<select name="selPC" class="select">
							<option value="0">�ֻ���</option>
							<%IF isArray(arrP) THEN
								For intP =0 To UBound(arrP,2)
								%>
							<option value="<%=arrP(0,intP)%>" <%IF Cstr(gpcode) = CStr(arrP(0,intP)) THEN%>selected<%END IF%>><%=arrP(1,intP)%></option>	
						<%  Next
							END IF%>	
							</select>
							<%ELSE%>
							<input type="hidden" name="selPC" value="<%=gpcode%>">
							<%=gpdesc%>
							<%END IF%>
							</td>
						</tr>
						<tr>
							<td width="100" align="center" bgcolor="<%= adminColor("tabletop") %>">�׷��</td>
							<td bgcolor="#FFFFFF"><input type="text" name="sGD" size="25" value="<%=db2html(gdesc)%>" maxlength="32" class="text"></td>
						</tr>		
						<tr>
							<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>
							<td bgcolor="#FFFFFF"><input type="text" size="2" name="sGS" id="sGS"  value="<%=gsort%>" class="text"></td>
						</tr> 
						<tr>
							<td align="center" bgcolor="<%= adminColor("tabletop") %>">���ÿ���</td>
							<td bgcolor="#FFFFFF"><input type="radio" name="eIsDisp" value="1" <%if gdisp then%>checked<%end if%>>Y <input type="radio" name="eIsDisp" value="0" <%if not gdisp then%>checked<%end if%>>N </td>
						</tr> 
					</table>
				</td> 
			</tr> 
			<tr>
				<td colspan="2" bgcolor="#FFFFFF" align="center" height="40"> 
					<input type="button" class="button" style="color:red;width:80px;" value="����" onClick="jsGroupSubmit();"> 
					<input type="button" class="button" style="color:blue;width:80px;" value="���" onClick="window.close();" > 
				</td>
			</tr>	
		</table>
		</form>	
	</td>
</tr>
<tr>
	<td><hr width="100%"></td>
</tr>
<tr>
	<td>
	    <p style="color:blue"> : ���콺 �巡�׷� �׷� ���������� �����մϴ�. <br>
	        &nbsp;&nbsp;���ϴ� ��ġ�� ���������� �� ��! <b>[��ü �׷��/���� ����]</b> �� �����ּ��� </p>
	    <%if sTarget <> "item" THEN%><p style="color:blue">: �׷� ���/���� �� [�� �׷�����]��ư�� �����ּ���  </p><%end if%>

	    <form name="frmL" method="post" action="eventgroup_process.asp">
	    <input type="hidden" name="eC" value="<%=eCode%>"> 
		<input type="hidden" name="mode" value="GS">  
		<input type="hidden" name="eCh" value ="<%=eChannel%>"> 
		<input type="hidden" name="sTarget" value ="<%=sTarget%>">    
		<input type="hidden" name="sGDarr" value =""> 
		<%IF isArray(arrList) THEN %> 
		    <div style="width:100%;text-align:right;padding:5px;">
		         <input  type="button" class="button"  style="font:bold;" value="��ü �׷��/���� ����" onClick="jsChSort();" >
		        <%IF eChannel ="M" then%>
		        <input type="button" class="button" value="���ñ׷� ��ġ��" onClick="jsSetJoin();">
		        <%END IF%>
		    </div>    
				<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				     
				<tr align="center">
				    
				    <%IF eChannel ="M" then%><td width="30" bgcolor="<%= adminColor("tabletop") %>">����</td>	<%END IF%>	
					<td bgcolor="<%= adminColor("tabletop") %>">�׷��ڵ�</td>					
					<td width="100" bgcolor="<%= adminColor("tabletop") %>">�����׷�</td>
					<td bgcolor="<%= adminColor("tabletop") %>">�׷��</td>
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">���ļ���</td>	 
					<td width="50" bgcolor="<%= adminColor("tabletop") %>">���ÿ���</td>	 
					<td width="140" bgcolor="<%= adminColor("tabletop") %>">����</td>
				</tr>
				<tbody id="subList">
				<%dim sumi,i ,eGCMoArr %>
				<%FOR intg = 0 To UBound(arrList,2) 
				    sumi = 0 
				    eGCMoArr = arrList(0,intg)
				%> 
				<tr <%if not arrList(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
				    <%IF eChannel ="M" then%><td  align="center"><input type="checkbox" name="chks" value="<%=arrList(0,intg)%>"></td><%END IF%>
					<td  >
					    <%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%>
					    
					    <% if intg < UBound(arrList,2) and eChannel ="M" then 
					    for i = 1 to (UBound(arrList,2)-intg)%>
					    <%if arrList(9,intg) = arrList(9,intg+i) then
					        sumi = sumi + 1 
					       
					        eGCMoArr = eGCMoArr &"," & arrList(0,intg+i)
					         %>
					    + <%=arrList(0,intg+i)%>
					    
					    <%else 
					     exit for
					    end if 
					    next
					end if
					    %> 
					    <input type="hidden" name="eGCMoArr" value="<%=eGCMoArr%>">
					    <input type="hidden" name="eGCArr" value="<%=arrList(0,intg)%>">
					    <input type="hidden" name="ePGCArr" value="<%=arrList(5,intg)%>">
					 </td>						
					<td  align="center">
					    <%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%>
					 </td>	
					<td  align="center"><input type="text" name="eGD" value="<%=db2html(arrList(1,intg))%>" size="25" class="text" maxlength="32"></td>	
					<td  align="center"><input type="text" name="eSArr" id="eSArr" value="<%=arrList(2,intg)%>" size="3"  style="text-align:center;"></td>		
					<td  align="center"><%if arrList(8,intg) then%>Y<%else%>N<%end if%>&nbsp; <input type="button" name="btnA" value="����" onclick="jsDispGroup('<%=arrList(0,intg)%>','<%if arrList(8,intg) then%>0<%else%>1<%END IF%>')"  class="button"></td>
					<td >
						<input type="button" name="btnU" value="����" onclick="jsSetGroup('<%=eCode%>','<%=arrList(0,intg)%>','<%=eChannel%>','<%=sTarget%>')" class="button">
						<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>','<%=eChannel%>','<%=sTarget%>')"  class="button">
						
						<%if sumi >0 and eChannel="M" then%><input type="button"  value="�׷쳪����" onclick="jsSetDivide('<%=arrList(0,intg)%>')"  class="button"> <%end if%>
					</td>					   									
				</tr>
			    <%   intg = intg+sumi
				NEXT%>
			    </tbody>
				</table>
		<%END IF%>	
	</form>
	</td>
</tr> 
<tr>
	<td align="center"><p> 
	    <%if sTarget <> "item" THEN%>
	    <input id="btn_g" type="button" class="button" style="height:30px; width:100px;" value="�� �׷�����"  > 
	    <%end if%>
	    </p> </td>
</tr> 
</table>       
<div id="divIpG" style="display:none;">
<%IF isArray(arrList) THEN %>
	<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
		<td>�׷��ڵ�</td>					
		<td>�����׷�</td>
		<td>�׷��</td>
		<td>���ļ���</td>					
		<td>�̹���</td>
		<td>���ÿ���</td>
		<td>����</td>
	</tr>
	<%FOR intg = 0 To UBound(arrList,2)%>				   						
	<tr <%if not arrList(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
		<td  ><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%></td>						
		<td  align="center"><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
		<td  align="center"><%=db2html(arrList(1,intg))%></td>	
		<td  align="center"><%=arrList(2,intg)%></td>									   									
		<td  align="center">    
			<a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a>  
		</td>					   								
		<td  align="center"><%if arrList(8,intg) then%>Y<%else%>N<%end if%></td>
		<td  align="center">
			<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrList(0,intg)%>','P')" class="button">
			<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">-->
			<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>','P')"  class="button">
			<% IF arrList(5,intg) = 0 THEN %>
			
			<% 		Response.Write "<a href='" & wwwUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			<% END IF %>
		</td>					   									
	</tr>
	<%NEXT%>
	</table>
<%END IF%>
</div>
<div id="divIpMG" style="display:none;">
<%IF isArray(arrList) THEN %>
	<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
		<td>�׷��ڵ�</td>					
		<td>�����׷�</td>
		<td>�׷��</td>
		<td>���ļ���</td>					
		<td>�̹���</td>
		<td>���ÿ���</td>
		<td>����</td>
	</tr>
	<%FOR intg = 0 To UBound(arrList,2)
	    sumi= 0
	%>				   						
	<tr <%if not arrList(8,intg) then%>bgcolor="gray"<%else%>bgcolor="#ffffff"<%end if%>>
		<td ><%IF arrList(5,intg) <> 0 THEN%><img src="/images/L.png">&nbsp;<%END IF%><%=arrList(0,intg)%>
		  <% if intg < UBound(arrList,2) and eChannel ="M" then 
				 for i = 1 to (UBound(arrList,2)-intg)%> 
				<%if arrList(9,intg) = arrList(9,intg+i) then
					sumi = sumi + 1 
				 %>
				 + <%=arrList(0,intg+i)%>
				<%else 
					exit for
				 end if 
				next
			end if    %>
		</td>						
		<td  align="center"><%IF isnull(arrList(7,intg))THEN%>�ֻ���<%ELSE%>[<%=arrList(5,intg)%>]<%=db2html(arrList(7,intg))%><%END IF%></td>	
		<td  align="center"><%=db2html(arrList(1,intg))%></td>	
		<td  align="center"><%=arrList(2,intg)%></td>									   									
		<td  align="center">    
			<a href="javascript:jsImgView('<%=arrList(3,intg)%>');"><img src="<%=arrList(3,intg)%>" width="50" border="0"></a>  
		</td>			
		<td  align="center"><%if arrList(8,intg) then%>Y<%else%>N<%end if%></td>		   									
		<td  align="center">
			<input type="button" name="btnU" value="����" onclick="jsGroupImg('<%=eCode%>','<%=arrList(0,intg)%>','M')" class="button">
			<!--<input type="button" name="btnD" value="����" onclick="jsDelGroup('<%=eCode%>','<%=arrList(0,intg)%>')"  class="button">-->
			<input type="button" name="btnD" value="��ǰ���" onclick="popRegItem('<%=eCode%>','<%=arrList(0,intg)%>','M')"  class="button">
			<% IF arrList(5,intg) = 0 THEN %>
			
			<% 		Response.Write "<a href='" & mobileUrl & "/event/eventmain.asp?eventid=" & eCode & "&eGC="& arrList(0,intg) &"' target='_blank'>�̸�����</a>"
			 %>
			<% END IF %>
		</td>					   									
	</tr>
	<%
	    intg = intg+sumi
	NEXT%>
	</table>
<%END IF%>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->