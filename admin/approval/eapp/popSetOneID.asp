<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���̵� ���
' History : 2014.01.02 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
dim oMember, arrList,intLoop
dim sUsername, oldpart_sn
Dim imgJoin,part_sn
 
sUsername =  requestCheckvar(Request("sUN"),32) 
'�������Ʈ ��������
	Set oMember = new CTenByTenMember
	oMember.Fusername 		= sUsername
	arrList = oMember.fnGetUserTreeList
	set oMember = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"> </script> 
<script type="text/javascript">
	//Ʈ����� Ŭ���� ���� ����
	function jsOpenClose(iValue){
		if(eval("document.all.divB"+iValue).style.display=="none"){
			eval("document.all.Fimg"+iValue).src = "/images/dtree/openfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tminus.png";
			eval("document.all.divB"+iValue).style.display=""; 
		}else{
			eval("document.all.Fimg"+iValue).src = "/images/dtree/closedfolder.png";
			eval("document.all.Timg"+iValue).src="/images/Tplus.png";
			eval("document.all.divB"+iValue).style.display="none"; 
		}
	}
	
	
	//�������
	function  jsSetId(iValue, sUid, sUname, ijobsn,sJobname){ 
	 if(document.all.divU.length== undefined ){  
		if (eval("document.all.divU").style.background =="white"){
				eval("document.all.divU").style.background = "yellow";
				if (document.frmS.hidUI.value==""){
					document.frmS.hidUI.value = sUid;
					document.frmS.hidUN.value = sUname;
					document.frmS.hidJS.value = ijobsn;
					document.frmS.hidJN.value = sJobname;
				}else{
					document.frmS.hidUI.value = document.frmS.hidUI.value +','+ sUid;
					document.frmS.hidUN.value = document.frmS.hidUN.value +','+ sUname;
					document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + ijobsn;
					document.frmS.hidJN.value = document.frmS.hidJN.value +','+ sJobname;
			 }
		}else{
			eval("document.all.divU").style.background = "white";
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(","); 
		var arrJN = document.frmS.hidJN.value.split(",");
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value = "";
		 document.frmS.hidJN.value ="";
	 		for(i=0;i<arrUI.length;i++){ 
	  			if(arrUI[i]!=sUid){
						if(document.frmS.hidUI.value==""){
							document.frmS.hidUI.value = arrUI[i];
							document.frmS.hidUN.value = arrUN[i];
							document.frmS.hidJS.value = arrJS[i];
							document.frmS.hidJN.value = arrJN[i];
						}else{
						document.frmS.hidUI.value = document.frmS.hidUI.value +','+ arrUI[i];
						document.frmS.hidUN.value = document.frmS.hidUN.value +','+ arrUN[i];
						document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + arrJS[i];
						document.frmS.hidJN.value = document.frmS.hidJN.value +','+ arrJN[i];
						}
	 			}
 			}
		}
	}else{
		if (eval("document.all.divU["+iValue+"]").style.background =="white"){
				eval("document.all.divU["+iValue+"]").style.background = "yellow";
				if (document.frmS.hidUI.value==""){
					document.frmS.hidUI.value = sUid;
					document.frmS.hidUN.value = sUname;
					document.frmS.hidJS.value = ijobsn;
					document.frmS.hidJN.value = sJobname;
				}else{
					document.frmS.hidUI.value = document.frmS.hidUI.value +','+ sUid;
					document.frmS.hidUN.value = document.frmS.hidUN.value +','+ sUname;
					document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + ijobsn;
					document.frmS.hidJN.value = document.frmS.hidJN.value +','+ sJobname;
			 }
		}else{
			eval("document.all.divU["+iValue+"]").style.background = "white";
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(","); 
		var arrJN = document.frmS.hidJN.value.split(",");
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value = "";
		 document.frmS.hidJN.value ="";
	 		for(i=0;i<arrUI.length;i++){ 
	  			if(arrUI[i]!=sUid){
						if(document.frmS.hidUI.value==""){
							document.frmS.hidUI.value = arrUI[i];
							document.frmS.hidUN.value = arrUN[i];
							document.frmS.hidJS.value = arrJS[i];
							document.frmS.hidJN.value = arrJN[i];
						}else{
						document.frmS.hidUI.value = document.frmS.hidUI.value +','+ arrUI[i];
						document.frmS.hidUN.value = document.frmS.hidUN.value +','+ arrUN[i];
						document.frmS.hidJS.value = document.frmS.hidJS.value + ',' + arrJS[i];
						document.frmS.hidJN.value = document.frmS.hidJN.value +','+ arrJN[i];
						}
	 			}
 			}
		}
	}		
}
	
	//���û�� ����(itype =1), ����(itype =2), ����(itype =3)�� �������
	function jsProcId(iType){
		var chkU;
		var arrUI = document.frmS.hidUI.value.split(",");
		var arrUN = document.frmS.hidUN.value.split(","); 
		var arrJS = document.frmS.hidJS.value.split(",");
		var arrJN = document.frmS.hidJN.value.split(",");
			
		if(iType==2){
			if(arrUI.length > 1 || $("#selAU option").size()>0){
				alert("���Ǵ� �Ѹ� ���ð����մϴ�.");
				return;
			}else{ 
				if(document.frmS.hidUI.value!=""){
					$("#selAU").append("<option value='"+document.frmS.hidUI.value+"-"+document.frmS.hidJS.value+"'>"+document.frmS.hidUN.value+" "+ document.frmS.hidJN.value+" ["+document.frmS.hidUI.value+"]</option>")
				}
			 }
		}else if(iType==3){
			for(i=0; i<arrUI.length;i++){
				chkU = 0;
				for(j=0; j<$("#selCU option").size();j++){ 
					if(($("#selCU option:eq("+j+")").val().split("-"))[0]==arrUI[i]){
					chkU = 1;
					}
				}
		 
				if (chkU ==0){
					if(arrUI[i]!=""){
					$("#selCU").append("<option value='"+arrUI[i]+"-"+arrJS[i]+"'>"+arrUN[i]+" "+arrJN[i]+" ["+arrUI[i]+"]</option>")
				}
				}
			}
		}else{
			
			for(i=0; i<arrUI.length;i++){
				chkU = 0;
				for(j=0; j<$("#selPU option").size();j++){ 
					if(($("#selPU option:eq("+j+")").val().split("-"))[0]==arrUI[i]){
					chkU = 1;
					}
				} 
				if (chkU ==0){
					if(arrUI[i]!=""){
					$("#selPU").append("<option value='"+arrUI[i]+"-"+arrJS[i]+"'>"+arrUN[i]+" "+arrJN[i]+" ["+arrUI[i]+"]</option>")
				}
				}
			}
		 }
		 
		 for(i=0;i<document.all.divU.length;i++){	
		 	document.all.divU[i].style.background = "white";
		 }	
		 document.frmS.hidUI.value ="";
		 document.frmS.hidUN.value =""; 
		 document.frmS.hidJS.value =""; 
		 document.frmS.hidJN.value ="";
	}
	
	//���û���
	function jsSelDel(iType){
		if (iType==2){
		 $("#selAU option:selected").remove();
		}else if(iType==3){
			$("#selCU option:selected").remove();
		}else{
			$("#selPU option:selected").remove();
		}
	}
	 

// ������ �ε�� ���� ���� ���缱 ��������
$(window).load(function(){  
	//���缱  
	 var arrAI = opener.document.frm.hidAI.value.split(",");
	 var arrAJ = opener.document.frm.hidAJ.value.split(",");
	 var arrATxt = opener.document.frm.hidATxt.value.split(","); 
	 for(i=0;i<arrAI.length;i++){
	 	if(arrAI[i]!=""){
	 	$("#selPU").append("<option value='"+arrAI[i]+"-"+arrAJ[i]+"'>"+arrATxt[i]+"</option>");
	}
	}
	
	//��������
	if(opener.document.frm.hidALI.value!=""){
	$("#selPU").append("<option value='"+opener.document.frm.hidALI.value+"-"+opener.document.frm.hidALJ.value+"'>"+opener.document.frm.hidALTxt.value+"</option>");  
}
	//���� 
	if(opener.document.frm.hidAI_H.value !=""){ 
	$("#selAU").append("<option value='"+opener.document.frm.hidAI_H.value+"'>"+opener.document.frm.hidATxt_H.value+"</option>");  
}
	//����
	 var arrRI = opener.document.frm.hidRfI.value.split(",") ;
	 var arrRTxt = opener.document.frm.sRfN.value.split(",") ;  
	 for(i=0;i<arrRI.length;i++){
	 	if(arrRI[i]!=""){
	 	$("#selCU").append("<option value='"+arrRI[i]+"'>"+arrRTxt[i]+"</option>");
	}
	}
	
});

//�˻� 
$(document).ready(function(){
	$("#btnSearch").click(function(){
		var username = escape($("#sUN").val()); 
		 var url="ajaxPartUserList.asp";
		 var params = "sUN="+username; 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){ 
		 		$("#divUL").html(args);
		 	},

		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 });
	}); 
	 
});
</script> 
<style>
	FORM {display:inline;}
</style>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"> 
	<tr>
		<td>���缱 ����<hr width="100%"></td>
	</tr>
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0"> 
				<tr>
					<td>
						<!-- ��� ����Ʈ-------------------->
						<form name="frmS" id="frm" method="post" onsubmit="return false;">
							<input type="hidden" name="hidUI" value="">
							<input type="hidden" name="hidUN" value=""> 
							<input type="hidden" name="hidJS" value=""> 
							<input type="hidden" name="hidJN" value=""> 
						<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
							<tr>
								<td>�����: <input type="text" name="sUN" id="sUN" value="<%=sUserName%>"> <input type="button" class="button" value="�˻�" id="btnSearch"></td>
							</tr>
							<tr>
								<td>
									<div id="divUL" style="padding:5px;border:gray solid 1px;width:260px;height:350px;overflow-y:auto;"> 
										<div></div> 
									<%IF isArray(arrList) THEN
											For intLoop = 0 To UBound(arrList,2) 
												part_sn = arrList(0,intLoop)
												
												'//�μ��� Ʋ������ �����̹����� Ʋ������
												if intLoop < UBound(arrList,2)  THEN 
													IF part_sn <> arrList(0,intLoop+1) THEN
														imgJoin = "joinbottom.gif"
													ELSE	
														imgJoin = "join.gif"
													END IF
												else
													imgJoin = "joinbottom.gif"
												end if	
												
												'//�μ��� Ʋ������ div ���� �ٲ۴�
												IF part_sn <> oldpart_sn THEN
													if intLoop <> 0 THEN
										%>
													</div>
												<%	end if%>
										<div id="divP<%=intLoop%>" style="cursor:hand;" onClick="jsOpenClose(<%=intLoop%>);"><img src="/images/<%IF susername ="" THEN%>Tplus<%ELSE%>Tminus<%END IF%>.png" align="absmiddle" id="Timg<%=intLoop%>"><img src="/images/dtree/<%IF susername ="" THEN%>closedfolder<%ELSE%>openfolder<%END IF%>.png" align="absmiddle" id="Fimg<%=intLoop%>"> <%=arrList(1,intLoop)%></div>  
										<div id="divB<%=intLoop%>" style="display:<%IF susername ="" THEN%>none<%END IF%>;cursor:hand;">	
									<%	END IF%>
										<div id="divU" style="padding:0 0 0 1;background:white;" onClick="jsSetId('<%=intLoop%>','<%=arrList(4,intLoop)%>','<%=arrList(5,intLoop)%>','<%=arrList(6,intLoop)%>','<%=arrList(8,intLoop)%>');"><img src="<%IF part_sn<>arrList(0,ubound(arrList,2)) then %>/images/dtree/line.gif<%else%>/images/blank.png<%END IF%>" align="absmiddle"><img src="/images/dtree/<%=imgJoin%>" align="absmiddle"> <%=arrList(5,intLoop)%>&nbsp;<%=arrList(8,intLoop)%> <font color="gray">[<%=arrList(4,intLoop)%>]</font> </div>
								<%		oldpart_sn = part_sn
										Next
									END IF%>
								</div>
								</td>
							</tr>
						</table>
					</form>
					<!-- //��� ����Ʈ-------------------->
					</td> 
					<td valign="top"> 
						<table width="200" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
							<!--- ���缱 ���� ����Ʈ-------------------->
							<tr>	
								<td>
									<input type="button" id="btnAdd" value="���� ��" style="color:blue;" class="button" onClick="jsProcId(1);" style="cursor:hand;">
								</td>
								<td>
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
											<tr>
												<td colspan="2"> + ���缱 (������� ��)</td>
											</tr>
											<tr>
												<td colspan="2"> 
														<select  name="selPU" id="selPU" multiple size="6" style="width:200px">
				
														</select> 
												</td>
											</tr>
											<tr>
												<td><img src="/images/mbtn_up.gif" align="absmiddle" id="sel_up" onClick="jsSelectMoveUp();"> <img src="/images/mbtn_down.gif" align="absmiddle" onClick="jsSelectMoveDown();"style="cursor:hand;"></td>
												<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle" onClick="jsSelDel(1)" style="cursor:hand;"></td>
											</tr>
										</table>
								</td>
							</tr>
							<!---  //���缱 ���� ����Ʈ-------------------->
							<!--- ���� ����Ʈ-------------------->
							<tr>
								<td>
									<input type="button" id="btnAdd" value="���� ��"   class="button" onClick="jsProcId(2);" style="cursor:hand;">
								</td>
								<td style="padding-top:10px;"> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0">
										<tr>
											<td>+ ����</td>
										</tr>	
										<tr>
											<td align="center">
												<select  name="selAU" id="selAU" multiple size="1" style="width:200px">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle"  onClick="jsSelDel(2)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>
								<!--- //���� ����Ʈ-------------------->
							<tr>
								<td colspan="2"><hr width="100%"></td>
							</tr>
								<!--- ���� ����Ʈ-------------------->
							<tr>
								<td>
									<input type="button" id="btnAdd" value="���� ��"  class="button" onClick="jsProcId(3);">
								</td>
								<td> 
										<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" border="0" >
										<tr>
											<td>+ ����</td>
										</tr>	
										<tr>
											<td align="center">
												<select  name="selCU" id="selCU" multiple size="4" style="width:200px">

												</select>
											</td>
										</tr>
											<tr>
													<td align="right"><img src="/images/mbtn_selectDel.gif" align="absmiddle" onClick="jsSelDel(3)" style="cursor:hand;"></td>
											</tr>
									</table>
								</td>
							</tr>	<!--- //���� ����Ʈ-------------------->
						</table>
					</td>
				</tr>
					<tr>
	<td align="center" colspan="3"><hr width="100%"><input type="button" class="button" value="���" onClick="jsSetReport();" style="cursor:hand;"></td>
</tr>
			</table>
		</td>
	</tr>

</table>
