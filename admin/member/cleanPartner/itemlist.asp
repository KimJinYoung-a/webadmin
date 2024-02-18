<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/cleanPartnerItem.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim yyyy , mm,d,mend,mdispend,mstart, monthDiff,favCnt
dim makerid,   cdl, cdm, cds
dim OnlySellyn, OnlyIsUsing,danjongyn,mwdiv,mode, dispCate
dim itemid, itemoption
dim arrList, intLoop
dim cCItem
dim iCurrpage, iPageSize,iTotCnt, iTotalPage,iPerCnt
dim chkImg
dim sSort

iCurrpage   = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
yyyy        = requestCheckVar(Request("yyyy1"),4)
mm          = requestCheckVar(Request("mm1"),2)
favCnt      = requestCheckVar(Request("favCnt"),10)
mwdiv       = requestCheckVar(Request("mwdiv"),2)
OnlySellyn  = requestCheckVar(Request("OnlySellyn"),2)
dispCate    = requestCheckvar(request("disp"),16)
makerid 	= requestCheckvar(Request("makerid"),32)
itemid      = requestCheckvar(Request("itemid"),255)
chkImg      = requestCheckvar(Request("chkImg"),1)
sSort				= requestCheckvar(Request("sSort"),2)
monthDiff				= requestCheckvar(Request("monthDiff"),6)
if (yyyy = "") then
	d = CStr(dateadd("m" ,-3, now()))
	yyyy = Left(d,4)
	mm = Mid(d,6,2)
end if

if (monthDiff = "") then
	monthDiff = "3"
end if 
  mend    = dateadd("m",1,yyyy&"-"&mm&"-01") 
    mdispend = dateadd("d",-1,dateadd("m",1,yyyy&"-"&mm&"-01"))
    mstart  =  dateadd("m", monthdiff*-1,mend)
if favCnt ="" then
    favCnt = 5
end if
   
if OnlySellyn ="" then OnlySellyn ="YS"
if mwdiv ="" then mwdiv ="U"
        
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 100		'�� �������� �������� ���� ��
    iPerCnt = 10		'�������� ������ ����
    
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
  	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

  
set cCItem = new CCleanItem
	cCItem.FCurrPage = iCurrpage		'����������
	cCItem.FPageSize = iPageSize	    '������������ 
	cCItem.FRectStdate   = mstart
  cCItem.FRectEddate   = mend
  cCItem.FRectWishCount=favcnt 
	cCItem.FRectMakerid	 = makerid
	cCItem.FRectDispCate = dispCate
	cCItem.FRectItemid   = itemid  
  cCItem.FRectSellYN   = OnlySellyn
  cCItem.FRectMwdiv    = mwdiv
  cCItem.FRectSort     = sSort
	arrList = cCItem.fnGetCleanItemList
	iTotCnt = cCItem.FTotCnt
set cCItem = nothing
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript"> 
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

function jsSetUseYN() {
	var frm = document.frmList;
	 
	if(typeof(frm.chkitem) =="undefined"){
	   return;
     }    
	 	
	if(!frm.chkitem.length){
	    if(!frm.chkitem.checked){
	 		alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 		return;
	 	}
	 	
	 	frm.itemidarr.value = frm.chkitem.value;
	 	 
	 }else{
    	for(i=0;i<frm.chkitem.length;i++){
    		if(frm.chkitem[i].checked) {
    	  			if (frm.itemidarr.value==""){
    	      			 frm.itemidarr.value =  frm.chkitem[i].value;
    	  			}else{
    	  	    		 frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
    	  			} 
    	  	} 
	  	 }

    	 if (frm.itemidarr.value == ""){
    	 	alert("������ ��ǰ�� �����ϴ�. ��ǰ�� ������ �ּ���");
	 		return;
	  	 }
	 }
	  

	if (confirm("���� ��ǰ�� ������ ó���Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}
 

function jsWinOpen(sUrl){ 
    var winOpen = window.open(sUrl,"popWin", "width=700 height=700 scrollbars=yes resizable=yes");
    winOpen.focus();
}


	 //����Ʈ ����
	 function jsSort(sValue,i){  
	  
	 	document.frm.sSort.value= sValue; 
	 	 
		   if (-1 < eval("document.all.img"+i).src.indexOf("_top")){
	        document.frm.sSort.value= sValue+"D";  
	    }else if (-1 < eval("document.all.img"+i).src.indexOf("_bot")){
	     		document.frm.sSort.value= sValue+"A";  
	    }else{
	       document.frm.sSort.value= sValue+"A";  
	    } 
	    
	   
		 document.frm.submit();
	}
 </script>
<form name="frm" method="get" action=""> 
<input type="hidden" name="menupos" value="<%= menupos %>"> 
<input type="hidden" name="sSort" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�����������: �˻��Ⱓ <% DrawYMBox yyyy, mm %>  ���� <input type="text" name="monthDiff" value="<%=monthDiff%>" size="4" style="text-align:right" class="input"> ���� �������� <span style="color:gray">[<%=mstart%>~<%=mdispend%>]</span> �Ǹż��� 0 ,
			 
		     &nbsp; &nbsp;
		     ���ü� <input type="text" name="favCnt" value="<%=favCnt%>" size="4"  style="text-align:right" class="input"> �� �̸� 
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	<tr    bgcolor="#FFFFFF" >	
		<td>
		    <table class="a" >
		        <tr>
		            <td> 
			            �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
            			&nbsp;
            			���� ī�װ�:  <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		            </td>
		            <td> 
			            &nbsp;��ǰ�ڵ� :
                    </td>
       	            <td   rowspan="2" >	
       	               <textarea rows="2" cols="10" name="itemid" id="itemid"><%=replace(itemid,",",chr(10))%></textarea>
		            </td>  
	            </tr>
	            <tr>
	                <td colspan="2">
	               	�Ǹ�:<% drawSelectBoxSellYN "OnlySellyn", OnlySellyn %>
                		&nbsp;
                	 <!--	���:<% drawSelectBoxUsingYN "OnlyIsUsing", OnlyIsUsing %>
                		&nbsp;-->
                	 <!--	����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
                		&nbsp;-->
                		�ŷ�����:<% drawSelectBoxMWU "mwdiv", mwdiv %>
                
                		
                	 </td>
                	</tr> 
			    </table>
		</td> 
	</tr> 
</table>
</form> 
<p> 
    <div align="left">
		<input type="button" class="button" value="���û�ǰ ������ó��" onClick="jsSetUseYN()"  >  
	</div>
    </p>
<!-- ����Ʈ ���� -->
<form name="frmList" method="post" action="procClean.asp">
    <input type="hidden" name="hidM" value="I">
	<input type="hidden" name="itemidarr" value="">   
	<input type="hidden" name="sRU" value="itemlist.asp?menupos=<%=menupos%>&makerid=<%=makerid%>&disp=<%=dispCate%>&OnlySellyn=<%=OnlySellyn%>&mwdiv=<%=mwdiv%>&itemid=<%=itemid%>&yyyy1=<%=yyyy%>&mm1=<%=mm%>&monthdiff=<%=monthdiff%>&favcnt=<%=favcnt%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b> <%=formatnumber(iTotCnt,0)%></b>
			&nbsp;
			������ : <b><%=formatnumber(iCurrpage,0)%>/ <%=formatnumber(iTotalPage,0)%></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><input type="checkbox" name="chkAll"  onClick="jsChkAll()"></td><!--list_lineup_bot_on-->
		<td onClick="javascript:jsSort('I','1');" style="cursor:hand;">��ǰ�ڵ� <img src="/images/list_lineup<%IF sSort="ID" THEN%>_bot_on<%ELSEIF sSort="IA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img1"></td>
		<td>�̹���</td>
		<td>�귣��</td>
		<td>��ǰ��</td> 
		<td>���ü�</td> 
		<td>�ǸŻ���</td>
		<td>�ŷ�����</td>
		<td onClick="javascript:jsSort('R','2');" style="cursor:hand;">����� <img src="/images/list_lineup<%IF sSort="RD" THEN%>_bot_on<%ELSEIF sSort="RA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img2"></td>
		<td onClick="javascript:jsSort('S','3');" style="cursor:hand;">�ǸŽ����� <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot_on<%ELSEIF sSort="SA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img3"></td>
	</tr>
	<%if isArray(arrList) then
	    For intLoop = 0 To UBound(arrList,2)
	    	'FItemList(i).Fimgsmall = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(FItemList(i).Fitemid) + "/" + rsget("smallimage")
	    %>
	<tr bgcolor="#FFFFFF" align="center">
	    <td><input type="checkbox" id="chkitem" name="chkitem" value="<%=arrList(0,intLoop)%>"></td>
	    <td><a href="<%=vwwwUrl%>/shopping/category_prd.asp?itemid=<%=arrList(0,intLoop)%>" target="_blank"><%=arrList(0,intLoop)%></a></td>
	    <td><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(arrList(0,intLoop))%>/<%=arrList(7,intLoop)%>"></td>
	    <td><a href="javascript:jsWinOpen('/admin/member/popbrandadminusing.asp?designer=<%=arrList(1,intLoop)%>');"><%=arrList(1,intLoop)%></td>
	    <td><%=arrList(2,intLoop)%></td>
	     <td><%=arrList(8,intLoop)%></td>
	    <td><%=fnColor(arrList(4,intLoop),"yn")%></td>
	    <td><%=mwdivName(arrList(3,intLoop))%></td> 
	    <td><%=arrList(5,intLoop)%></td>
	    <td><%=arrList(6,intLoop)%></td>
	</tr>
    <% Next 
    else
    %>
    <tr bgcolor="#FFFFFF" >
        <td colspan="10" align="center">�˻����ǿ� �ش��ϴ�  ������ �������� �ʽ��ϴ�.</td>
    </tr>
    <%
	end if%>
</table>
</form>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table> 
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->