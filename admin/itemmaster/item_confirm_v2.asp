<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���δ���ǰ �󼼸���Ʈ
' History : 2014.01.06 ������ ����
'						currstate: 0-���ιݷ�,1-���δ��,2-���κ���,5-���δ��(���û),7-���οϷ�,9-��ü���
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"-->
<%
Dim sListType, sCurrState, sSort, sMode
Dim dispCate, makerid, itemname, itemcount, itemid
Dim iCurrpage,ipagesize,iTotCnt,iTotalPage
Dim clsWait,arrList, intLoop
Dim clsPartner,idefaultmargine
dim onlyNotSet
dim cdl, cdm, cds, ctrState

	sListType =  requestCheckVar(request("sLT"),1)
	sCurrstate =  requestCheckVar(request("sCS"),1)
	sSort =  requestCheckVar(request("sS"),2)
	dispCate = requestCheckvar(request("disp"),16)
	makerid	= requestCheckvar(Request("makerid"),32)
	itemname	= requestCheckvar(Request("itemname"),64)
	itemid= requestCheckvar(Request("itemid"),255)
	onlyNotSet =  requestCheckVar(request("onlyNotSet"),1)

	cdl = requestCheckvar(request("cdl"),10)
	cdm = requestCheckvar(request("cdm"),10)
	cds = requestCheckvar(request("cds"),10)

	ctrState= requestCheckvar(request("selCtr"),1)
	
 	if sCurrState = "" THEN sCurrState = "1"
 	if sSort = "" THEN sSort = "ID"
 		
  if dispcate <> "" and makerid <> "" then
  	iPageSize = 25
  else		
 		iPageSize = 50
	end if

	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1

	if itemid<>"" then
		dim iA ,arrTemp,arrItemid
		itemid = replace(itemid,chr(13),"")
		arrTemp = Split(itemid,chr(10))

		iA = 0
		do while iA <= ubound(arrTemp)

			if trim(arrTemp(iA))<>"" then
				'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.04;������)
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

	if (onlyNotSet = "Y") then
		dispCate = ""
	end if

set clsWait = new CWaitItemlist2014

	if (onlyNotSet = "Y") then
		clsWait.Fcatecode 	= "n"
	else
		clsWait.Fcatecode 	= dispCate
	end if

	clsWait.FRectCate_Large   = cdl
	clsWait.FRectCate_Mid     = cdm
	clsWait.FRectCate_Small   = cds

	clsWait.Fmakerid		= makerid
	clsWait.Fitemname		= itemname
	clsWait.Fcurrstate		= sCurrstate
	clsWait.FSort			= sSort
	clsWait.FPageSize		= iPageSize
	clsWait.FCurrPage		= iCurrPage
	clsWait.Fitemid			= itemid
	clsWait.FRectctrState	= ctrState
	arrList = clsWait.fnGetWaitItemList
	iTotCnt	= clsWait.FTotCnt
 set clsWait = nothing
'  if dispCate ="n" then dispCate = ""
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
dim dctrState	
	'�˻����ǿ� �귣�� ���� ��� ��ǥ��ǰ������ Ȯ��(�ϰ����ν� üũ����)
	dctrState = 7 '���� ���� ��� ��ǥ�����´� ���Ϸ��.. 
	if makerid <>"" then
		if isArray(arrList) then
		dctrState = arrList(16,0)
		end if
	end if
%>
<style>
	#dialog {display:none; position:absolute;left:100;top:100; z-index:9100;background:#efefef; padding:10px;width:650;}
	#mask {display:none; position:absolute; left:0; top:0; z-index:9000; background:url(http://webadmin.10x10.co.kr/images/mask_bg.png) left top repeat;}
	#boxes .window {position:fixed; left:0; top:0; display:none; z-index:99999;}
</style>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript">

//�˻�
function jsSearch(sValue){
	if(sValue!=""){
		document.frm.sCS.value = sValue;
	}
	document.frm.submit();
}

//����Ʈ ����
function jsSort(sValue,i){
	 	document.frm.sS.value= sValue;

		   if (-1 < eval("document.frmList.img"+i).src.indexOf("_alpha")){
	        document.frm.sS.value= sValue+"D";
	    }else if (-1 < eval("document.frmList.img"+i).src.indexOf("_bot")){
	     		document.frm.sS.value= sValue+"A";
	    }else{
	       document.frm.sS.value= sValue+"D";
	    }
		 document.frm.submit();
	}

// ������ �̵�
function jsGoPage(iCP)
	{
		document.frm.iCP.value=iCP;
		document.frm.submit();
	}

//-----------------------------------------------
//���º���
	function jsUniWaitState(itemid){
	var ret = confirm('���δ��� �����Ͻðڽ��ϱ�?');

	if (ret){
		 document.frmList.hidM.value="U";
		 document.frmList.itemid.value = itemid;
		 document.frmList.sCS.value =5;
		 document.frmList.action ="doitemregboru.asp";
		 document.frmList.submit();
	}
}

var chkCnt = 0 ;
 //���� ���û�ǰ ���º���
function jsMultiWaitState(currstate){
	var frm = document.frmList;
	 var itemcount = 0;
	 var count2 = 0;
	if(typeof(frm.chkitem) !="undefined"){
	 	if(!frm.chkitem.length){
	 		if(!frm.chkitem.checked){
	 			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	 		}
	 		 frm.itemidarr.value = frm.chkitem.value;
	 		 itemcount = 1;
	 		 if (frm.hidC2.value==2){
	 		 count2 = 1;
	 		}
	  }else{
	  	for(i=0;i<frm.chkitem.length;i++){
	  		if(frm.chkitem[i].checked) {
	  			if (frm.itemidarr.value==""){
	  			 frm.itemidarr.value =  frm.chkitem[i].value;
	  			}else{
	  			 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
	  			}
	  			 itemcount = itemcount+ 1;
	  			 if (frm.hidC2[i].value==2){
					 		 count2 = 1;
					 		}
	  		}

	  	}

	  	if (frm.itemidarr.value == ""){
	  		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	  	}
	  }
	}else{
		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
		return;
	}

 if (currstate ==1){
 		  if ( chkCnt > 0 ){
 	  	alert("���� �����͸� ó�����Դϴ�.��� �� �ٽ� ����ó�� ���ּ���");
 	  	return;
 	  }
 	  
 	  var dCtrState = "<%=dctrState%>";
 	   if (dCtrState!="7"){
 	  	alert("���̿Ϸ�� �귣��� ������ �Ұ����մϴ�.\n���Ȯ�� �� ó�����ּ���");
 	  	return;
 	  }
 	  
 		if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?\n��ü��� ��ǰ�� ��� ����Ʈ�� �ٷ� ����Ǹ�, \n�ٹ����ٹ�ۻ�ǰ�� �԰� �Ϸ� �� ��ǰ�� ���µ˴ϴ�.")){
				  document.itemArrreg.itemid.value = frm.itemidarr.value ; 
				  chkCnt ++;
				  $("#btnSubmit").prop("disabled", true);
				   document.itemArrreg.submit();
		}
}else{
		if(count2>0&&currstate==2){
			alert("�����Ͻ� ��ǰ �߿� 3�� �������� �ֽ��ϴ�.\n�ش��ϴ� ��ǰ�� ���κ���(���Ͽ�û)�� �Ͽ��� ���ιݷ�(���ϺҰ�) ó���ǹǷ� ���� ��Ź �帳�ϴ�.");
		}

		frm.itemcount.value = itemcount;
			//	var popWin = window.open("item_confirm_pop.asp?sCS="+currstate+"&itemcount="+itemcount,"popW","width=600,height=500");
		 $("#dv2").hide();
		 $("#dv0").hide();
 	 	 $('html, body').animate({scrollTop:0});

		var maskHeight = $(document).height();
		var maskWidth = $(document).width();
		$('#mask').css({'width':maskWidth,'height':maskHeight});
		 $('#boxes').show();
		$('#mask').show();
	//	var winH = $(window).height();
	//	var winW = $(document).width();
	//	$("#dialog").css('top', winH/2-$("#dialog").height()/2);
	//	$("#dialog").css('left', winW/2-$("#dialog").width()/2);
		$("#dialog").show();
		$("#dv"+currstate).show();
	}	

 }


	$('#mask').click(function () {
		$('#boxes').hide();
		$('.window').hide();
		$('#dialog').hide();
	});


  function jsCancel(){
  	document.frmList.sMsgcd.value= "";
 		document.frmList.sMsg.value = "";
 		document.frmList.itemcount.value="";
 		document.frmList.itemidarr.value="";
  	 $( "#dialog" ).hide();
  	 $('#mask').hide();
  	 $('#boxes').hide();
  }


 //���ΰź�ó��
 function jsConfirm(currstate){
 	var chkCount = 0;
 	var iMsgcd = "";
 	var sMsg = "";
 	var iNo = ""; 
 	for(i=0;i<eval("document.all.chkV"+currstate).length;i++){ 
 		if(eval("document.all.chkV"+currstate)[i].checked){
 		chkCount = chkCount + 1;
 		iNo = eval("document.all.chkV"+currstate)[i].value;
 		if (iMsgcd==""){
 			iMsgcd = eval("document.all.chkV"+currstate)[i].value; 
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 				sMsg = eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = $("#sp"+currstate+iNo).text();
 			}
 		}else{
 		    iMsgcd = iMsgcd +"^"+ eval("document.all.chkV"+currstate)[i].value;  
 			if (eval("document.all.chkV"+currstate)[i].value==999){
 					sMsg = sMsg +"^"+eval("document.all.sM"+currstate).value;
 			}else{
 				sMsg = sMsg +"^"+ $("#sp"+currstate+iNo).text();
 			}
 		}
 	}
 	}
 	if(chkCount == 0){
 		alert("���� �ź� ������ �Ѱ� �̻� �������ּ���");
 		return;
 	}
 	
  
 	document.frmList.sMsgcd.value= iMsgcd;
 	document.frmList.sMsg.value = sMsg;
 	document.frmList.hidM.value = "M";
 	document.frmList.sCS.value = currstate;
  document.frmList.submit();
}


 //����ó��
 function jsApproval(itemid,makerid,ctrstate){
 	 
 	  if (ctrstate!="7"){
 	  	alert("���̿Ϸ�� �귣��� ������ �Ұ����մϴ�.\n���Ȯ�� �� ó�����ּ���");
 	  	return;
 	  }
 	
 	  if ( chkCnt > 0 ){
 	  	alert("���� �����͸� ó�����Դϴ�.��� �� �ٽ� ����ó�� ���ּ���");
 	  	return;
 	  }

			if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?\n��ü��� ��ǰ�� ��� ����Ʈ�� �ٷ� ����Ǹ�, \n�ٹ����ٹ�ۻ�ǰ�� �԰� �Ϸ� �� ��ǰ�� ���µ˴ϴ�.")){
				  document.itemreg.itemid.value = itemid;
				  document.itemreg.makerid.value = makerid;
				  chkCnt ++;
				   document.itemreg.submit();
				}
}


//�󼼳��� ����
	function popItemModify(itemid,designer){
	var popwin = window.open('wait_item_modify.asp?itemid=' + itemid + '&designer=' + designer,'waititemmodify','width=1000,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//��ü ����
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
	   	   	if(frm.chkitem.disabled==false){
		   	 	frm.chkitem.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					 	if(frm.chkitem[i].disabled==false){
					frm.chkitem[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

//�������� ���̾�ǥ��
$(document).ready(function(){
 $("div.dlog").click(function(){
 	var divindex =$("div.dlog").index(this);
 	var itemid =$(this).attr("id") ;
 	var url="item_confirm_ajaxLog.asp";
		 var params = "itemid="+itemid;
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("div.dsub").empty().hide();
		 		$("div.dsub").eq(divindex).show();
		 		$("div.dsub").eq(divindex).html(args);
		 	},
		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 //	 alert(e.responseText);
		 	}
	})
	})

	$("div.dlog").mouseleave(function(){
		$("div.dsub").empty().hide();
		})
});

function ViewItemDetail(itemno){
	window.open('/designer/itemmaster/viewitem.asp?itemid='+itemno ,'window1','');
}

function jsPopOption(itemno){
 var winOpt = window.open("/common/pop_upchewaititemoptionedit.asp?itemid="+itemno,"editItemOption","width=800,height=400,scrollbars=yes,resizable=yes"); 
 winOpt.focus();
}
 </script>
<form name="itemreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/doWaitItemToMultiReg_byadmin.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="makerid" value="">
	<input type="hidden" name="sCS" value="<%=sCurrstate%>"> 
</form>
<form name="itemArrreg" method="post" action="<%= ItemUploadUrl %>/linkweb/items/doWaitItemToMultiOneReg_byadmin.asp" onsubmit="return false;" enctype="multipart/form-data">
	<input type="hidden" name="itemid" value="">  
</form>
<table width="100%" border="0" cellpadding="1" cellspacing="0" class="a">
	<tr>
		<td><!-- �˻�---------------------------------->
			<form name="frm" method="get" action="">
			<input type="hidden" name="iCP" value="1">
			<input type="hidden" name="menupos" value="<%= request("menupos") %>">
			<input type="hidden" name="sS" value="<%=sSort%>"><!--����-->
			<input type="hidden" name="sLT" value="<%=sListType%>"><!--����ƮŸ��(b:�귣��, c:ī�װ�)-->
			<input type="hidden" name="sCS" value="<%=sCurrstate%>">
				<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" class="a">
					<tr align="center" bgcolor="#FFFFFF">
						<td rowspan="2" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
						<td  bgcolor="#FFFFFF" align="left">
							<table border="0" cellpadding="3" cellspacing="0" class="a">
								<tr>
									<td>�귣��: <%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
									<td> ��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="20"></td>
									<td>�ӽ��ڵ�:</td>
									<td rowspan="2">
							 			<textarea rows="3" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
									</td>
								</tr>
								<tr>
									<td colspan="3">
										���� ī�װ�:  <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
										&nbsp;
										����<!-- #include virtual="/common/module/categoryselectbox.asp"-->
									</td>
								</tr>
								<tr>
									<td colspan="3">
										<input type="checkbox" name="onlyNotSet" value="Y" <% if (onlyNotSet = "Y") then %>checked<% end if %> > ����ī�װ� ������ ��ǰ��
										&nbsp;&nbsp;
										������: 
										<select name="selCtr" class="select">
											<option value="">��ü</option>
											<option value="Y" <%IF ctrState ="Y" then%>selected<%END IF%>>���Ϸ�</option>
											<option value="N" <%IF ctrState ="N" then%>selected<%END IF%>>���̿Ϸ�</option>
										</select>
									</td>
								</tr>
							</table>
						</td>
						<td rowspan="2"  width="50" bgcolor="#EEEEEE">
							<input type="button" class="button_s" value="�˻�" onClick="jsSearch('');">
						</td>
					</tr>
				</table>
			</form>
		</td><!-- //�˻�---------------------------------->
	</tr>
	<tr>
		<td>
			<div style="padding:5px"></div>
		</td>
	</tr>
	<tr>
		<td><!-- action ---------------------------------->
				<table width="100%" border="0" cellpadding="5" cellspacing="1"  class="a">
					<tr>
						<td> + �귣��, ����ī�װ� ���ý� [�ϰ�����] ��ư�� Ȱ��ȭ �˴ϴ�. ���� ���� ��ǰ ó���� �ӵ��� ������ �� ������ ��ٷ��ּ���. 
		 <!--input type="button" class="button" value="���¿���"--></td>
						<td align="right">
							<input type="button" class="button" value="���κ���(���Ͽ�û)" onClick="jsMultiWaitState(2);">
							<input type="button" class="button" value="���ιݷ�(���ϺҰ�)" onClick="jsMultiWaitState(0);">
							<%if dispcate <> "" and makerid <> ""  then%>
							<input type="button" class="button" id="btnSubmit" value="  �ϰ�����  " onClick="jsMultiWaitState(1);" style="background-color:#F29661;">
							<%end if%>
						</td>
					</tr>
				</table>
		</td><!-- //action ---------------------------------->
	</tr>
	<tr>
		<td><!-- List ---------------------------------->
			<form name="frmList" method="post" action="doitemregboru.asp">
			<input type="hidden" name="hidM" value="">
			<input type="hidden" name="itemidarr" value="">
			<input type="hidden" name="itemid" value="">
			<input type="hidden" name="itemcount" value="">
			<input type="hidden" name="sCS" value="">
			<input type="hidden" name="sMsgcd" value="">
			<input type="hidden" name="sMsg" value="">
			<input type="hidden" name="sS" value="<%=sSort%>">
			<input type="hidden" name="sRU" value="item_confirm.asp?sLT=<%=sListType%>&makerid=<%=makerid%>&disp=<%=dispCate%>&sCS=<%=sCurrstate%>&sS=<%=sSort%>">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
				<tr bgcolor="#FFFFFF">
					<td colspan="15" height="25" align="left">�˻����: <b><%=iTotCnt%></b> &nbsp; ������: <b><%=iCurrpage%>/<%=iTotalPage%></b></td>
				</tr>
				<tr class="a" height="25" bgcolor="#DDDDFF" align="center">
					<td><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
					<td width="80" onClick="javascript:jsSort('I','7');" style="cursor:hand;">�ӽ��ڵ� <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot<%ELSEIF sSort="IA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img7"></td>
					<td>�̹���</td> 
					<td width="90" onClick="javascript:jsSort('B','1');" style="cursor:hand;">�귣��ID <img src="/images/list_lineup<%IF sSort="BD" THEN%>_bot<%ELSEIF sSort="BA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
					<td onClick="javascript:jsSort('N','2');" style="cursor:hand;">��ǰ�� <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
					<td width="60" onClick="javascript:jsSort('S','3');" style="cursor:hand;">�ǸŰ� <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot<%ELSEIF sSort="SA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
					<td width="60" onClick="javascript:jsSort('A','4');" style="cursor:hand;">���԰� <img src="/images/list_lineup<%IF sSort="AD" THEN%>_bot<%ELSEIF sSort="AA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td>
					<td>�ɼ� �߰�����</td>    
					<td>�ŷ�����</td>
					<td width="40" onClick="javascript:jsSort('M','5');" style="cursor:hand;">���� <img src="/images/list_lineup<%IF sSort="MD" THEN%>_bot<%ELSEIF sSort="MA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
					<td>����</td>
					<td>����ī�װ� <font color="blue">(+�߰�ī�װ�)</font></td>
					<td width="160" onClick="javascript:jsSort('L','6');" style="cursor:hand;">�������� <img src="/images/list_lineup<%IF sSort="LD" THEN%>_bot<%ELSEIF sSort="LA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img6"></td>
					<td width="100">
						<select name="selCS" class="select" onChange="jsSearch(this.value);">
							<%sbOptItemWaitStatus sCurrState%>
						</select>
					</td>
				</tr>
				<%IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					%>
				<tr bgcolor="<%if arrList(16,intLoop) <> 7 THEN%>#DDDDFF<%else%>#FFFFFF<%end if%>" align="center">
				 	<td><input type="checkbox" name="chkitem" value="<%= arrList(0,intLoop) %>" <%IF arrList(7,intLoop) <> 1 and arrList(7,intLoop) <> 5  THEN%>disabled<%END IF%>>
				 		<input type="hidden" name="hidC2" value="<%=arrList(9,intLoop)%>">
				 		</td>
					<td><a href="javascript:popItemModify('<% =arrList(0,intLoop) %>','<%=arrList(2,intLoop) %>')"><%=arrList(0,intLoop)%></a>
						<br/> 
						<%if arrList(16,intLoop) <> 7 THEN%>
					 	[���̿Ϸ�]
						<%end if%>
						</td>
					<td><%IF arrList(1,intLoop) <> "" THEN
						dim imgsubdir
						imgsubdir = GetImageSubFolderByItemid(arrList(0,intLoop))
						%>
						<img src="<%=partnerUrl%>/waitimage/basic/<%=imgsubdir%>/<%= arrList(1,intLoop)%>" width="50" height="50">
						<%END IF%>
					</td> 
					<td><%=arrList(2,intLoop)%></td>
					<td>
						<%=arrList(3,intLoop)%> 
						<a href="javascript:ViewItemDetail('<%=arrList(0,intLoop)%>');"><font color="blue">[�̸�����]</font></a>
						<%
							Dim keyword, chk
							keyword = arrList(3,intLoop)
							If InStr(keyword, "_") > 0 Then 
								chk = InStr(keyword, "_") - 1
								keyword = Mid(keyword, 1, chk)
							End If
							keyword = URLEncodeUTF8(keyword)

							Response.Write "<a href='http://shopping.naver.com/search/all.nhn?query="& keyword &"&pagingIndex=1&pagingSize=40&viewType=list&sort=rel' target='"& arrList(0,intLoop) &"'><font color='blue'>[������ Ȯ���ϱ�]</font></a>"
						%>
						
					</td>
					<td width="60" align="right"><%=formatnumber(arrList(5,intLoop),0)%>&nbsp;</td>
					<td width="60" align="right"><%=formatnumber(arrList(4,intLoop),0)%>&nbsp;</td>
					<td><a href="javascript:jsPopOption('<%= arrList(0,intLoop) %>');"><%if arrList(20,intLoop) >0 then%><font color=red>Y</font><%else%>N<%end if%></a></td>
					<td><%IF arrList(11,intLoop) <> arrList(12,intLoop) THEN%><font color="red"><%end if%><%=mwdivName(arrList(12,intLoop))%></td>
					<td width="40" align="right"><%IF arrList(6,intLoop) <> arrList(10,intLoop)  THEN%><font color=red><%END IF%><%=arrList(6,intLoop)%>%&nbsp;</td>
					<td><% if arrList(15,intLoop)="Y" then %>
						<font color=red>����</font><%=arrList(13,intLoop)-arrList(14,intLoop) %>
						<% end if %>
					</td>
					<td align="left"><a href="javascript:popItemModify('<% =arrList(0,intLoop) %>','<%=arrList(2,intLoop) %>')">
						<% if Not isNull(arrList(18,intLoop)) then Response.write replace(arrList(18,intLoop),"^^",">") %> &nbsp;<%if arrList(19,intLoop)  > 0 then%><font color="blue"><%end if%>(+<%=arrList(19,intLoop)%>)</a></td>
					<td width="160"><div id="<%= arrList(0,intLoop) %>" class="dlog" style="cursor:hand;" ><%=arrList(8,intLoop)%></div>
						<div style="position:relative;background-color:#eeeeee">
						 <div id="dLogSub" class="dsub" style="position:absolute;left:-80px;top:0px;z-index:100;background-color:white;"></div>
					 </div>
						</td>
					<td><font color="<%=GetCurrStateColor(arrList(7,intLoop))%>"><%=GetCurrStateName(arrList(7,intLoop))%></font>
							<% if (arrList(7,intLoop)="2") or (arrList(7,intLoop)="0") then %>
							<span style="line-height:23px;"><a href="javascript:jsUniWaitState('<%=arrList(0,intLoop) %>')"><br><font color="#000000">[���δ�⺯��]</font></a></span>
							<% elseif  (arrList(7,intLoop)="1") or (arrList(7,intLoop)="5") then%>
						 	&nbsp;<input id="btnApp" name="btnApp" type="button" class="button" value="������" style="color:blue;" onclick="jsApproval('<%=arrList(0,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(16,intLoop)%>')">
							<% end if %>

					</td>
				</tr>
				<%	Next
					ELSE
				%>
				<tr bgcolor="#ffffff">
					<td align="center" colspan="14">��ϵ� ������ �����ϴ�.</td>
				</tr>
				<%
				END IF%>
			</table>
</form>
		</td><!-- //List ---------------------------------->
	</tr>
	<!-- ������ ���� -->
<% 	Dim iStartPage,iEndPage,iX,iPerCnt
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
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
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
	</td>
	</tr>
</table>
<div id="boxes">
<div id="mask"></div>
<div id="dialog">
<!-- #include virtual="/admin/itemmaster/item_confirm_inc.asp"-->
</div>
</div>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
