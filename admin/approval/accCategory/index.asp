<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ����  ����Ʈ
' History : 2011.03.09 ������  ����
'########################################################### 
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/accCategoryCls.asp" -->
<%
Dim clsAcc 
Dim arrList, intLoop
Dim ipcateidx, icateidx,sACCUSECD,sACCNM,sisNoSet 
Dim iCurrPage, iPageSize,iTotCnt,iTotalPage
Dim iCpcateidx,iCcateidx

ipcateidx = requestCheckvar(Request("selCL"),10)
IF ipcateidx = "" THEN ipcateidx = 0
icateidx 	= requestCheckvar(Request("iCS"),10)
IF icateidx = "" THEN icateidx = 0
sACCUSECD = requestCheckvar(Request("sAUCD"),15)
sACCNM 		= requestCheckvar(Request("sANM"),50)
sisNoSet 	= requestCheckvar(Request("chkNS"),1)
iCurrPage = requestCheckvar(Request("iCP"),10)
IF iCurrPage = "" THEN iCurrPage = 1
iPageSize = 30

Set clsAcc = new CAccCategory
	clsAcc.FACCPCateIdx =  ipcateidx
	clsAcc.FACCCateIdx  =  icateidx 
	clsAcc.FACCUSECD    =  sACCUSECD   
	clsAcc.FACCNM       =  sACCNM      
	clsAcc.FisNoSet     =  sisNoSet    
	clsAcc.FCurrPage     =  iCurrPage    
	clsAcc.FPageSize     =  iPageSize    
	arrList = clsAcc.fnGetACCCDList
	iTotCnt	= clsAcc.FTotCnt
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>  
<script  type="text/javascript">
<!--
// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}
	   
//���� ����
function jsMngCategory(){
	var winC = window.open("categoryList.asp","popC","width=600, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
} 
 
//�������� �Ⱥб��ذ���
function jsSetDivide(acccd){
	var winC = window.open("accDivide.asp?Acccd="+acccd,"popC","width=600, height=800, resizable=yes, scrollbars=yes");
	winC.focus();
} 
//�˻�
function jsSearch(){
		document.frm.iCS.value = $("#selC").val();  
		document.frm.submit();
}
//���� ���� =========================================================================================================
$(document).ready(function(){
	$("#selCL").change(function(){
		var iValue = $("#selCL").val(); 
		var url="/admin/approval/accCategory/ajaxCate.asp";
		 var params = "sVar=selC&selCL="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divC").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
 
	$("#selCCL").change(function(){
		var iValue = $("#selCCL").val(); 
		var url="/admin/approval/accCategory/ajaxCate.asp";
		 var params = "sVar=selCC&selCL="+iValue;  
		   
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divCC").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});
//���� ����
function jsSetCategory(){
	if($("#selCCL").val() ==0){
		alert("������� �������ּ���");
		return;
	}
	if($("#selCC").val() ==0){
		alert("�߰����� �������ּ���");
		return;
	}
	
	var ischecked =false;
    
    for (var i=0;i<frmReg.elements.length;i++){
		//check optioon
		var e = frmReg.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}
	
	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}
	
	if (confirm('����Ͻðڽ��ϱ�?')){  
			frmReg.iccidx.value = $("#selCC").val();
 	    frmReg.submit();
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

	function checkThis(comp){
    AnCheckClick(comp)
} 

//ī�װ�����
function jsDelCate(iValue){
	if(confirm("�����Ͻ� ������ ������ �����Ͻðڽ��ϱ�?")){
		document.frmDel.hidCDIdx.value = iValue;
		document.frmDel.submit();
	}
}
//-->
</script>
<form name="frmDel" method="post" action="procCategory.asp">
	<input type="hidden" name="hidM" value="D">
	<input type="hidden" name="hidCDIdx" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="iCS" value="<%=icateidx%>">
	<input type="hidden" name="selCL" value="<%=ipcateidx%>">
</form>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a"> 
<tr>
	<td><form name="frm" method="get" action="index.asp" style="margin:0px;">
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iCS" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					 ����:
				 	<select name="selCL"  id="selCL">
					<option value="0">--����--</option>
					<%clsAcc.sbGetOptAccCategory 1,0,ipcateidx %>			
					</select> 
					>
					<span id="divC">
					<select name="selC"  id="selC">
					<option value="0">--����--</option>
					<% IF ipcateidx > 0 THEN
					clsAcc.sbGetOptAccCategory 2,ipcateidx,icateidx 
						END IF
					%>			
					</select> 
					</span> 
					&nbsp;&nbsp;
					���������: <input type="text" name="sANm" value="<%=sACCNm%>" size="20">
					&nbsp;&nbsp;
					���������ȣ: <input type="text" name="sAUCD" value="<%=sACCUseCD%>" size="15">
					&nbsp;&nbsp;
					<input type="checkbox" name="chkNS" value="Y" <%IF sIsNoSet ="Y" THEN%>checked<%END IF%>> ���� ������ ������
				</td>
				<td  width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr> 
		</table>
	</form>
	</td>
</tr>  
<tr>
	<td><hr width="100%"></td>
</tr>
<form name="frmReg" method="post" action="procCategory.asp" style="margin:0px;">
	<input type="hidden" name="hidM" value="S">
	<input type="hidden" name="iccidx" value=""> 
	<input type="hidden" name="menupos" value="<%= menupos %>">
<tr>
	<td>	<input type="button" class="button" value="���� ����" onClick="jsMngCategory();">   
		 &nbsp;|&nbsp; 
	 ����: 
				 	<select name="selCCL"  id="selCCL">
					<option value="0">--����--</option>
					<%clsAcc.sbGetOptAccCategory 1,0,iCpcateidx %>			
					</select> 
					>
					<span id="divCC">
					<select name="selCC"  id="selCC">
					<option value="0">--����--</option>
					<% IF iCpcateidx > 0 THEN
					clsAcc.sbGetOptAccCategory 2,iCpcateidx,iCcateidx 
						END IF
					%>			
					</select> 
					</span> 
					<input type="button" class="button" value="���� ���" onClick="jsSetCategory();"> : ����  ���� ���ð��� ���
	</td>
</tr>
<%Set clsAcc = nothing%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<tr>
	<td> 
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="15">
					�˻���� : <b><%=iTotCnt%></b> &nbsp;
					������ : <b><%= iCurrPage %> / <%=iTotalPage%></b>
				</td>
			</tr>
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>"> 
				<td>����</td>
				<td>����</td>
				<td>���������ڵ�</td> 
				<td>���������</td> 
				<td>����ó</td> 
				<td>�Ⱥб���</td> 
				<td>ó��</td> 
			</tr>
			<%  
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2) 
				%>
			<tr height=30 align="center" bgcolor="#FFFFFF">	
				<td><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" <%IF not isNull(arrList(3,intLoop)) THEN %>disabled<%END IF%>></td>
				<td><%IF not isNull(arrList(3,intLoop)) THEN %><%=arrList(6,intLoop)%> > <%=arrList(4,intLoop)%><%END IF%></td>			
				<td><%=arrList(1,intLoop)%></td>	
				<td><%=arrList(2,intLoop)%></td>	
				<td><%if arrList(8,intLoop) then %>10x10<%end if%>&nbsp;
					<%if arrList(9,intLoop) then %>����<%end if%>
				</td> 
				<td><%=arrList(10,intLoop)%></td>
				<td>
				
				<input type="button" value="�Ⱥб��ذ���" class="button" onClick="jsSetDivide(<%=arrList(0,intLoop)%>);"
				<%IF   isNull(arrList(3,intLoop)) THEN %>disabled<%END IF%>
				>
				&nbsp;
				
				<input type="button" value="����" class="button" onClick="jsDelCate(<%=arrList(7,intLoop)%>);">
				</td>	 
			</tr>
		<%	Next
			ELSE%>
			<tr height=5 align="center" bgcolor="#FFFFFF">				
				<td colspan="4">��ϵ� ������ �����ϴ�.</td>	
			</tr>
			<%END IF%>
		</table>	 
	</td> 
</tr> 
</form>	
<!-- ������ ���� -->
		<%
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>        
					    </tr>        
					</table>
				</td>
			</tr>
</table>
<!-- ������ �� -->
</body>
</html>
 



	