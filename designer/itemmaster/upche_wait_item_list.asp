<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/itemregcls.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2014.asp"--> 
<% 
Dim owaititem,ix,page,itemname, i
Dim dispCate,sCurrState,sSort,upchemanagecode
dispCate = requestCheckvar(request("disp"),16)
sCurrstate =  requestCheckVar(request("sCS"),1)
upchemanagecode = requestCheckVar(request("upchemanagecode"),32)
if sCurrstate = "" THEN sCurrstate = "A"
page = requestCheckVar(request("page"),10)
if (page="") then page=1
''itemname = requestCheckVar(request("itemname"),64) ''�÷��̿��� ��û ' ġȯ ����
itemname = LEFT(trim(request("itemname")),64)
itemname = replace(itemname,"--","")

sSort =  requestCheckVar(request("sS"),2)

set owaititem = new CWaitItemlist
owaititem.FPageSize = 20
owaititem.FCurrPage = page
owaititem.FRectDesignerID = session("ssBctID")
owaititem.FRectitemname = itemname
owaititem.Fcatecode = dispCate
owaititem.FRectCurrState = sCurrstate
owaititem.FRectSort = sSort
owaititem.FRectUpchemanagecode = upchemanagecode
owaititem.WaitProductList

%> 
<style>
FORM {display:inline;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
function ViewItemDetail(itemno){
	window.open('/designer/itemmaster/viewitem.asp?itemid='+itemno ,'window1','width=960,height=600,scrollbars=yes,status=no');
}
function TnSearchItem(sValue){
	document.frm.page.value = "";
		if(sValue!=""){
		document.frm.sCS.value = sValue;
	}
	document.frm.submit();
}

function ChangeOrderMakerFrame(){ 
	var frm = document.frmBuyPrc;
	var upfrm = document.frmArrupdate; 
	var itemcount = 0;
	if(typeof(frm.cksel) !="undefined"){
	 	if(!frm.cksel.length){
	 		if(!frm.cksel.checked){
	 			alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	 		}
	 		 upfrm.itemid.value = frm.cksel.value;
	 		 itemcount = 1;
	  }else{
	  	for(i=0;i<frm.cksel.length;i++){
	  		if(frm.cksel[i].checked) {	   	    			
	  			if (upfrm.itemid.value==""){
	  			upfrm.itemid.value =  frm.cksel[i].value;
	  			}else{
	  			upfrm.itemid.value =upfrm.itemid.value+ "|" +frm.cksel[i].value;
	  			} 
	  			 itemcount = itemcount+ 1;
	  		}	 
	  	}
	  } 	
	  	if (upfrm.itemid.value == ""){
	  		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 			return;
	  	} 
	}else{
		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
		return;
	}  

	var ret = confirm('���� ��ǰ�� �����Ͻðڽ��ϱ�?');
	if (ret){
	 upfrm.submit();
		}  
}
</script>
<script>
// ============================================================================
// �ɼǼ���
function PopUpcheItemOptionEdit(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_upchewaititemoptionedit.asp?' + param ,'PopUpcheItemOptionEdit','width=700,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ============================================================================
// �̹�������
function PopUpcheItemImageEdit(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemimage.asp?' + param ,'PopUpcheItemImageEdit','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

//��ü ����
function jsChkAll(){	
var frm;
frm = document.frmBuyPrc;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.cksel) !="undefined"){
	   	   if(!frm.cksel.length){
	   	   	if(frm.cksel.disabled==false){
		   	 	frm.cksel.checked = true;	  
		   	} 	 
		   }else{
				for(i=0;i<frm.cksel.length;i++){
					 	if(frm.cksel[i].disabled==false){
					frm.cksel[i].checked = true;
				}
			 	}		
		   }	
	   }	
	} else {	  
	  if(typeof(frm.cksel) !="undefined"){
	  	if(!frm.cksel.length){
	   	 	frm.cksel.checked = false;	  
	   	}else{
			for(i=0;i<frm.cksel.length;i++){
				frm.cksel[i].checked = false;
			}	
		}		
	  }	
	
	}
	
} 

//����Ʈ ����
function jsSort(sValue,i){ 
	 	document.frm.sS.value= sValue;
	 	 
		   if (-1 < eval("document.frmBuyPrc.img"+i).src.indexOf("_alpha")){
	        document.frm.sS.value= sValue+"D";  
	    }else if (-1 < eval("document.frmBuyPrc.img"+i).src.indexOf("_bot")){
	     		document.frm.sS.value= sValue+"A";  
	    }else{
	       document.frm.sS.value= sValue+"D";  
	    } 
		 document.frm.submit();
	} 
	
	
//�������� ���̾�ǥ��
$(document).ready(function(){
 $("div.dlog").click(function(){
 	var divindex =$("div.dlog").index(this);
 	var itemid =$(this).attr("id") ;
 	var url="item_confirm_ajaxLog.asp";
		 var params = "hidM=D&itemid="+itemid; 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){  
		 		$("div.dLsub").empty().hide();
		 		$("div.dLsub").eq(divindex).show();
		 		$("div.dLsub").eq(divindex).html(args);
		 	}, 
		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���1");
		 //	 alert(e.responseText);
		 	} 
	}) 
	})
	
	$("div.dlog").mouseleave(function(){ 
		$("div.dLsub").empty().hide();
		})
		
	$("div.dState").click(function(){
 	var divindex =$("div.dState").index(this);
 	var itemid =$(this).attr("id") ;
 	var url="item_confirm_ajaxLog.asp";
		 var params = "hidM=S&itemid="+itemid;  
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){  
		 		$("div.dSsub").empty().hide();
		 		$("div.dSsub").eq(divindex).show();
		 		$("div.dSsub").eq(divindex).html(args);
		 	}, 
		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���2");
		 //	 alert(e.responseText);
		 	} 
	}) 
	})
	
	$("div.dState").mouseleave(function(){ 
		$("div.dSsub").empty().hide();
		})
});
	  
</script>


<!-- ǥ ��ܹ� ����--> 
 	<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="sCS" value=""> 
	<input type="hidden" name="sS" value=""> 
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>"> 
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td colspan="3" background="/images/tbl_blue_round_02.gif"></td> 
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="30" >
        <td rowspan="2" background="/images/tbl_blue_round_04.gif"></td>
        <td colspan="3" valign="top">
         ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox_upche.asp"-->&nbsp;
    	   ��ǰ��:<input type="text" name="itemname" size="20" value="<%= itemname %>">&nbsp;&nbsp;
    	   ��ü��ǰ�ڵ�:<input type="text" name="upchemanagecode" size="20" value="<%= upchemanagecode %>">&nbsp;
    	   <a href="javascript:TnSearchItem('')"><img src="/admin/images/search2.gif" width="74" height="22" align="absmiddle" border="0"></a>
    	   <hr width="100%">
        </td>
         <td rowspan="3" background="/images/tbl_blue_round_05.gif"></td>
     </tr>
     <tr> 
     		<td>	�˻���� : �� <font color="red"><% = owaititem.FTotalCount %></font>��</td> 
     		<td> <font color="blue">+ ��������, ���λ��� ����Ʈ�� ���콺�� Ŭ���ϸ� �� �α׸� Ȯ�� �Ͻ� �� �ֽ��ϴ�.</font></td> 
        <td height="30" valign="top" align="right"> 
        	<input type="button" class="button" value="���û�ǰ����" onClick="ChangeOrderMakerFrame()">
        </td> 
    </tr>
</table> 
</form> 
<form name="frmBuyPrc" method="post"> 
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="30"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
			<td onClick="javascript:jsSort('T','1');" style="cursor:hand;"><b>�ӽ�</b>�ڵ� <img src="/images/list_lineup<%IF sSort="TD" THEN%>_bot<%ELSEIF sSort="TA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img1"></td>
			<td width="80" onClick="javascript:jsSort('U','2');" style="cursor:hand;">��ü�ڵ� <img src="/images/list_lineup<%IF sSort="UD" THEN%>_bot<%ELSEIF sSort="UA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img2"></td>
			<td onClick="javascript:jsSort('N','3');" style="cursor:hand;">��ǰ�� <img src="/images/list_lineup<%IF sSort="ND" THEN%>_bot<%ELSEIF sSort="NA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img3"></td>
			<td width="60" onClick="javascript:jsSort('S','4');" style="cursor:hand;">�ǸŰ� <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot<%ELSEIF sSort="SA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img4"></td> 
			<td  onClick="javascript:jsSort('L','5');" style="cursor:hand;">�������� <img src="/images/list_lineup<%IF sSort="LD" THEN%>_bot<%ELSEIF sSort="LA" THEN%>_top<%ELSE%>_alpha<%END IF%>.png" id="img5"></td>
			<td width="60">
				<select name="selCS" class="select" onChange="TnSearchItem(this.value);">
				<%sbOptItemWaitStatus sCurrState%>
				</select></td>
			<td width="50">�ɼ�</td>
	    </tr>
<% if owaititem.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="9" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<% else %>
	<% for ix=0 to owaititem.FresultCount-1 %> 
		<tr class="a" height="25" bgcolor="#FFFFFF">
			<td align="center"  width="30"> 
			<input type="checkbox" name="cksel" value="<% =owaititem.FItemList(ix).Fitemid %>" 	<% If (owaititem.FItemList(ix).FCurrState = 7) then %>disabled<% End if %>>
		  </td>
			<td align="center"><%= owaititem.FItemList(ix).Fitemid %></td>
			<td align="center"><%= owaititem.FItemList(ix).Fupchemanagecode %></td>
			<% if owaititem.FItemList(ix).FCurrState="7" then %>
			<td align="left">&nbsp;<% =owaititem.FItemList(ix).Fitemname %>&nbsp;&nbsp;<a href="http://www.10x10.co.kr/street/designershop.asp?itemid=<% =owaititem.FItemList(ix).Flinkitemid %>" target="_blank"><font color="blue">(����)</font></a></td>
			<% else %>
			<td align="left"><a href="upche_wait_item_modify.asp?itemid=<% =owaititem.FItemList(ix).Fitemid %>&menupos=<%= menupos %>"><% =owaititem.FItemList(ix).Fitemname %></a>&nbsp;&nbsp;<a href="javascript:ViewItemDetail('<% =owaititem.FItemList(ix).Fitemid %>')"><font color="blue">(�̸�����)</font></a></td>
			<% end if %>
			<td align="center"><%= FormatNumber(owaititem.FItemList(ix).Fsellcash,0) %></td> 
			<td align="center"><div id="<%=owaititem.FItemList(ix).Fitemid%>" class="dlog" style="cursor:hand;" ><% =owaititem.FItemList(ix).Flastupdate  %></div>
						<div style="position:relative;background-color:#eeeeee"> 
						 <div id="dLogSub" class="dLsub" style="position:absolute;left:-80px;top:0px;z-index:100;background-color:white;"></div>
					 </div>  </td>
			<td align="center"> 
				<div id="<%=owaititem.FItemList(ix).Fitemid%>" class="dState" style="cursor:hand;" ><font color="<%=GetCurrStateColor(owaititem.FItemList(ix).FCurrState) %>"><%=GetCurrStateName(owaititem.FItemList(ix).FCurrState)%></font></div>
				<div style="position:relative;background-color:#eeeeee"> 
						 <div id="dStateSub" class="dSsub" style="position:absolute;left:-120px;top:0px;z-index:100;background-color:white;"></div>
					 </div>
		 	</td>
			<td align="center">
            <% if (owaititem.FItemList(ix).FCurrState <> "7") then %>
				<a href="javascript:PopUpcheItemOptionEdit('<%= owaititem.FItemList(ix).Fitemid %>')">
				<img src="/images/icon_modify.gif" border="0" align="absbottom">
				</a>
            <% end if %>
			</td>
		</tr> 
    <% next %>
<% end if %>
	</form> 
<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
		<% if owaititem.HasPreScroll then %>
			<a href="javascript:NextPage('<%= owaititem.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + owaititem.StartScrollPage to owaititem.StartScrollPage + owaititem.FScrollCount - 1 %>
			<% if (ix > owaititem.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(owaititem.FCurrPage) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if owaititem.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
 
<form name="frmArrupdate" method="post" action="delwaititemarr.asp">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="itemid" value="">
</form>
 
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->