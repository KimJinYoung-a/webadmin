<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/partners/cleanPartnerBrand.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim brandusing, partnerusing, rdoutbrnad, research
dim makerid, groupid
dim yyyy1,mm1, newbrandgbn
dim nowdate, mode, mduserid, catecode
dim outlevel
dim dispcate ,standardCateCode
dim monthdiff, yyyy, mm, d,favcnt,mend,mstart
dim sSort
dim iCurrpage, iPageSize,iTotCnt, iTotalPage,iPerCnt

iCurrpage   = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
brandusing  = requestCheckvar(request("brandusing"),10)
partnerusing= requestCheckvar(request("partnerusing"),10)
rdoutbrnad= requestCheckvar(request("rdoutbrnad"),2)

research = requestCheckvar(request("research"),10)
catecode = requestCheckvar(request("catecode"),20)
mduserid = requestCheckvar(request("mduserid"),32)
outlevel = requestCheckvar(request("outlevel"),10)
makerid  = requestCheckvar(request("makerid"),32)
groupid  = requestCheckvar(request("groupid"),10)
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
newbrandgbn = requestCheckvar(request("newbrandgbn"),10)
standardCateCode    = requestCheckvar(request("standardCateCode"),16)
sSort				= requestCheckvar(Request("sSort"),2)
yyyy = requestCheckvar(request("yyyy"),10)
mm  = requestCheckvar(request("mm"),10)
if sSort = "" then sSort ="BA"
if research="" and brandusing="" then brandusing="Y"
if research="" and rdoutbrnad="" then rdoutbrnad="N"
if research="" and rdoutbrnad="" then newbrandgbn="N"
if yyyy1="" then
	nowdate = CStr(dateadd("m",-1,now()))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
end if


if (yyyy = "") then
	d = CStr(dateadd("m" ,-3, now()))
	yyyy = Left(d,4)
	mm = Mid(d,6,2)
end if
 
 mend    = dateadd("d",-1,dateadd("m",1,yyyy1&"-"&mm1&"-01"))
 mstart  = yyyy&"-"&mm&"-01"
if favCnt ="" then
    favCnt = 10
end if

	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 500		'�� �������� �������� ���� ��
    iPerCnt = 10		'�������� ������ ����
    
dim opartner
set opartner = new CPartnerUser
opartner.FCurrPage = iCurrpage		'����������  
opartner.FPageSize = iPageSize
opartner.FRectDesignerID = makerid
opartner.FRectYYYYMM = yyyy1 + "-" + mm1
opartner.FRectisusing = brandusing
opartner.FRectPartnerIsusing = partnerusing
opartner.FRectnewbrandgbn = newbrandgbn
opartner.FRectGroupid = groupid
opartner.FRectStdate   = mstart
opartner.FRectEddate   = mend
opartner.FRectWishCount=favcnt 
opartner.FRectDispCate = standardCateCode
opartner.FRectSort     = sSort	
'opartner.FRectMdUserID = mduserid
'opartner.FRectCatecode = catecode
'opartner.FRectmakerlevel = outlevel

opartner.FRectOutReqBrand = rdoutbrnad

opartner.GetOutBrandList

iTotCnt = opartner.FTotalCount 
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script type="text/javascript">
function MakeOutBrand(){
	if (confirm('�ۼ��Ͻðڽ��ϱ�?')){
	    document.actfrm.mode.value="makeoutbrand";
	    document.actfrm.target="_self";
		document.actfrm.submit();
	}
}

function PopUpcheInfo(v){
	//window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
}

function popSearchGroupID(frmname,compname){
    var popwin = window.open("/admin/member/popupcheselect.asp?frmname=" + frmname + "&compname=" + compname,"popSearchGroupID","width=800 height=680 scrollbars=yes resizable=yes");
    popwin.focus();
}

function brnadOutProc(imakerid){
    if (confirm(imakerid+' �귣�带 ���� �Ͻðڽ��ϱ�?')){
        document.actfrm.makerid.value=imakerid;
        document.actfrm.mode.value="prcoutbrand";
        document.actfrm.target="prcoutbrand";
		document.actfrm.submit();
    }
}

function scmKillProc(imakerid){
    if (confirm(imakerid+' �귣�带 SCM �α��� ���� �Ͻðڽ��ϱ�?')){
        document.actfrm.makerid.value=imakerid;
        document.actfrm.mode.value="prcscmnotusing";
        document.actfrm.target="prcoutbrand";
		document.actfrm.submit();
    }
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
	
	
	function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkmid) !="undefined"){
	   	   if(!frm.chkmid.length){
	   	   	if(frm.chkmid.disabled==false){
		   	 	frm.chkmid.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkmid.length;i++){
					 	if(frm.chkmid[i].disabled==false){
					frm.chkmid[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkmid) !="undefined"){
	  	if(!frm.chkmid.length){
	   	 	frm.chkmid.checked = false;
	   	}else{
			for(i=0;i<frm.chkmid.length;i++){
				frm.chkmid[i].checked = false;
			}
		}
	  }

	}

}

function jsSetUseYN() {
	var frm = document.frmList;
	 
	if(typeof(frm.chkmid) =="undefined"){
	   return;
     }    
	 	
	if(!frm.chkmid.length){
	    if(!frm.chkmid.checked){
	 		alert("������ �귣�尡 �����ϴ�. �귣�带 ������ �ּ���");
	 		return;
	 	}
	 	
	 	frm.makeridarr.value = frm.chkmid.value;
	 	 
	 }else{
    	for(i=0;i<frm.chkmid.length;i++){
    		if(frm.chkmid[i].checked) {
    	  			if (frm.makeridarr.value==""){
    	      			 frm.makeridarr.value =  frm.chkmid[i].value;
    	  			}else{
    	  	    		 frm.makeridarr.value = frm.makeridarr.value + "," +frm.chkmid[i].value;
    	  			} 
    	  	} 
	  	 }

    	 if (frm.makeridarr.value == ""){
    	 	alert("������ �귣�尡 �����ϴ�. �귣�带 ������ �ּ���");
	 		return;
	  	 }
	 }
	  

	if (confirm("���� �귣�带 ������ ó���Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function popItemSellEdit(designerid,usingyn){
	var popwin = window.open('/admin/shopmaster/itemviewset.asp?menupos=24&makerid=' + designerid + '&usingyn=' + usingyn  ,'popItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popOffItemSellEdit(designerid,itemgubun,usingyn){
	var popwin = window.open('/admin/offshop/shopitemlist.asp?menupos=184&research=on&page=1&ckonlyusing=on&designer=' + designerid + '&itemgubun=' + itemgubun + '&usingyn=' + usingyn ,'popOffItemSellEdit','width=1000,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

</script>

<% if (C_ADMIN_AUTH) and (opartner.FResultCount>0) then %>
<font color=red >�����ڸ޴�</font> : <input type="button" class="button_auth" value="�� �����ϱ�" onClick="MakeOutBrand()">
<br>
<% end if %>

<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="sSort" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">  
	<tr align="center" bgcolor="F4F4F4">
	    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
        <td bgcolor="#FFFFFF" align="left">
        	���س�� <% DrawYMBox yyyy1,mm1 %>
        	&nbsp;&nbsp;
        	�귣��ID : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
        	&nbsp;&nbsp;
        	��ü(�׷��ڵ�) : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="12" >
			<input type="button" class="button" value="Code�˻�" onclick="popSearchGroupID(this.form.name,'groupid');" >&nbsp;&nbsp;
        	&nbsp;&nbsp;
        	�űԺ귣�屸��
        	<select name="newbrandgbn">
        	<option value="">��ü
        	<option value="N" <%=CHKIIF(newbrandgbn="N","selected","")%> >�űԺ귣��(����� 6���� �̳�)
        	<option value="O" <%=CHKIIF(newbrandgbn="O","selected","")%> >�űԺ귣�� ����
        	</select>
        	&nbsp;&nbsp;
        	�귣���뿩��
        	<select name="brandusing">
        	<option value="">��ü
        	<option value="Y" <%=CHKIIF(brandusing="Y","selected","")%> >�����
        	<option value="N" <%=CHKIIF(brandusing="N","selected","")%> >������
        	</select>
        	&nbsp;&nbsp;
        	SCM���¿���
        	<select name="partnerusing">
        	<option value="">��ü
        	<option value="Y" <%=CHKIIF(partnerusing="Y","selected","")%> >Y
        	<option value="N" <%=CHKIIF(partnerusing="N","selected","")%> >N
        	</select>
        	<div style="padding:5 0 0 0;">	���� ī�װ�: <%= fnStandardDispCateSelectBox(1,"", "standardCateCode", standardCateCode, "")%></div>
        	<!--
			&nbsp;&nbsp;
			����� : <% drawSelectBoxCoWorker "mduserid", mduserid %>
			&nbsp;&nbsp;
			��ü���� : <% SelectBoxBrandCategory "catecode", catecode %>
			&nbsp;&nbsp;
			�ܰ豸�� :
			<select name="outlevel" >
			<option value="">��ü
			<option value="5" <% if outlevel="5" then response.write "selected" %> >level-5
			<option value="3" <% if outlevel="3" then response.write "selected" %> >level-3
			<option value="0" <% if outlevel="0" then response.write "selected" %> >level-0
			</select>
            -->
        </td>
        <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
    		<input type="button" class="button_s" value="�˻�" onClick="document.frm.submit();">
    	</td>
	</tr> 
	<tr>
	    <td bgcolor="#FFFFFF" align="left">
	        <input type="radio" name="rdoutbrnad" value="N" <%=CHKIIF(rdoutbrnad="N","checked","")%> >�� �˻�����<br>
	        <input type="radio" name="rdoutbrnad" value="YY" <%=CHKIIF(rdoutbrnad="YY","checked","")%> >������� �귣�� 
	        (�����Ǹſ�ON 12�������� & �����Ǹſ�OF 12�������� & �ǸŻ�ǰ��<strong>[ON]</strong> 0 & �Ż�ǰ ��� 0) <!-- & �űԺ귣������--><br>
	        <input type="radio" name="rdoutbrnad" value="YM" <%=CHKIIF(rdoutbrnad="YM","checked","")%> >������� �귣�� 
	        (�����Ǹſ�ON <% DrawYMBoxdynamic "yyyy",yyyy,"mm", mm,"" %>  �� ���� <font color="gray">[<%=mstart%>~<%=mend%>]</font> �Ǹż��� 0
	        , ���ü� <input type="text" name="favCnt" value="<%=favCnt%>" size="4"  style="text-align:right" class="input"> �� �̸� 
	        ) <!-- & �űԺ귣������--><br>
	    </td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��--> 
	</form>
	<br/>
<p> 
    <div>
			+ �˻����ǿ��� �������귣�� ������ �׼ǿ��� <font color="blue">1.��ǰ����</font> ��ư�� ���� ��ǰ ��뿩�� �� �Ǹſ��� 'N'���� ���� �� <font color="red">2. �귣������</font>��ư�� ������ �귣�� ������ ó���� ���ּ���
		</div>
</p>
 
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class=a>
<tr>
    <td colspan="25" bgcolor="#FFFFFF" height="30"> 
    		�˻���� : <b> <%=formatnumber(iTotCnt,0)%></b>
			&nbsp;
			������ : <b><%=formatnumber(iCurrpage,0)%>/ <%=formatnumber(iTotalPage,0)%></b> 
   </td>
</tr>
<tr align="center"  bgcolor="<%= adminColor("tabletop") %>">
	 
	<td width="80" rowspan=2  onClick="javascript:jsSort('B','1');" style="cursor:hand;">�귣��ID  <img src="/images/list_lineup<%IF sSort="BD" THEN%>_bot_on<%ELSEIF sSort="BA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img1"></td>
	<td rowspan=2>�귣���</td>
	<td width="50" rowspan=2>�׷��ڵ�</td>
	<td rowspan=2>����ڸ�</td>
	<td width="80" rowspan=2>����ڹ�ȣ</td>
	<%if rdoutbrnad ="YM" then%>
	<td width="50" rowspan=2>���ü�</td>
	<%end if%>
	<td width="70" rowspan=2 onClick="javascript:jsSort('R','2');" style="cursor:hand;">������<br>(�귣������) <img src="/images/list_lineup<%IF sSort="RD" THEN%>_bot_on<%ELSEIF sSort="RA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img2"></td>
	<td width="60" rowspan=2>3������<br>�Ż�ǰ��<br>[�ۼ���]</td>
	<td width="100" colspan=2>�����Ǹſ�<br>[�ۼ���]</td>
	<td width="150" colspan=3>����ǰ��<br>[�ۼ���]</td>
	<td width="50" >�Ǹ�<br>��ǰ��</td>
	<td width="50" colspan=3>[ON]����������<br>��ǰ��</td>
	
	<td width="100" colspan=3>�귣���뿩��</td>
<!--	<td width="100" colspan=2>��Ʈ��Ʈ���¿���</td>
	<td width="60" rowspan=2>Ŀ�´�Ƽ</td>-->
	<td width="50" rowspan=2>scm<BR>���¿���</td>
	<td width="60" rowspan=2>�����α���<br>(�귣��)<br>[����]</td>
	<td width="60" rowspan=2>�����α���<br>(�׷����)<br>[�ۼ���/1�Ⱓ]</td>
	<td width="100" rowspan=2>�׼�</td>
	<!--
	<td width="40" rowspan=2>level</td>
	<td width="70" rowspan=2>3������<br>�����</td>
	<td width="70" rowspan=2>�⺻����</td>
	-->
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td width="60"  onClick="javascript:jsSort('S','3');" style="cursor:hand;">ON  <img src="/images/list_lineup<%IF sSort="SD" THEN%>_bot_on<%ELSEIF sSort="SA" THEN%>_top_on<%ELSE%>_top<%END IF%>.png" id="img3"></td>
	<td width="60">OFF</td>
	<td width="50">ON</td>
	<td width="50">OFF</td>
	<td width="50">ETC</td>
	<td width="50">ON</td>
	<td width="50">��ü</td>
	<td width="50">��Ź</td>
	<td width="50">����</td>
	<td width="40">ON</td>
	<td width="40">OFF</td>
	<td width="40">���޸�</td>
	<!--<td width="50">�ٹ�����</td>
	<td width="50">���޸�</td>-->
</tr>





<% if opartner.FResultCount<1 then %>
	<tr bgcolor="#FFFFFF" >
		<td colspan="25" align=center>[�����Ͱ� �����ϴ�.]
		<% if rdoutbrnad="N" and makerid="" and groupid="" then %>
		 <input type="button" value="�����ϱ�" onClick="MakeOutBrand()"> 
		<% end if %> 
		</td>
	</tr>
<% else %>
	<% for i=0 to opartner.FResultCount -1 %>
	<% if opartner.FPartnerList(i).Fisusing="Y"	then %>
	<tr bgcolor="#FFFFFF">
	<% else %>
	<tr bgcolor="#EEEEEE">
	<% end if %>
	   
		<td><a href="javascript:PopBrandInfoEdit('<%= opartner.FPartnerList(i).Fmakerid %>')"><%= opartner.FPartnerList(i).Fmakerid %></a></td>
		<td><%= opartner.FPartnerList(i).Fmakername %></td>
		<td><%= opartner.FPartnerList(i).Fgroupid %></td>
		<td><%= opartner.FPartnerList(i).Fcompany_name %></td>
		<td><%= opartner.FPartnerList(i).Fcompany_no %></td>
			<%if rdoutbrnad ="YM" then%>
		<td align="center"><%= opartner.FPartnerList(i).FfavCount %></td>
		<%end if%>
		<td align=center><%= Left(opartner.FPartnerList(i).Fbrandregdate,10) %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).Fnewitemcount,0) %></td>
		<td align=center><%= opartner.FPartnerList(i).FlastsellDateON %></td>
		<td align=center><%= opartner.FPartnerList(i).FlastsellDateOF %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).Fcurrentusingitemcnt,0) %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).Foffcurrentusingitemcnt-opartner.FPartnerList(i).Fetccurrentusingitemcnt,0) %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).Fetccurrentusingitemcnt,0) %></td>
		
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).Fcurrentsellitemcnt,0) %></td>
		
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).FUCount,0) %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).FWCount,0) %></td>
		<td align=center><%= FormatNumber(opartner.FPartnerList(i).FMCount,0) %></td>
		
		<td align=center>
			<% if opartner.FPartnerList(i).Fisusing="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>
		<td align=center>
			<% if opartner.FPartnerList(i).Fisoffusing="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>
		<td align=center>
			<% if opartner.FPartnerList(i).Fisextusing="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>
	<!--	<td align=center>
			<% if opartner.FPartnerList(i).Fstreetusing="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>
		<td align=center>
			<% if opartner.FPartnerList(i).Fextstreetusing="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>
		<td align=center>
			<% if opartner.FPartnerList(i).Fspecialbrand="Y" then %>
			O
			<% else %>
			X
			<% end if %>
		</td>-->
		<td align=center>
		    <a href="javascript:PopBrandAdminUsingChange('<%= opartner.FPartnerList(i).Fmakerid %>')">
		    <% if (opartner.FPartnerList(i).Fisusing="N") and (opartner.FPartnerList(i).Fpartnerusing="Y") then %>
		    <font color="red"><b>O</b></font>
		    <% else %>
		        <% if opartner.FPartnerList(i).Fpartnerusing="Y" then %>
		        O
		        <% else %>
		        X
		        <% end if %>
		    <% end if %>
		    </a>
		</td>
		<td align=center><%=opartner.FPartnerList(i).FLastPartnerLogindate%></td>
		<td align=center><%=opartner.FPartnerList(i).Flastgrouplogindate%></td>
		<td align=center>
			<%if rdoutbrnad <>"N" then%> 
			<div style="padding:3 0 3 0;">
			 <input type="button" value="1.��ǰ����" onClick="popItemSellEdit('<%= opartner.FPartnerList(i).Fmakerid %>','Y');" class="button" <%IF opartner.FPartnerList(i).Fcurrentusingitemcnt <= 0 then %>disabled<%else%>style="color:blue;"<%end if%> > >>
			</div> 
				<div style="padding:3 0 3 0;">
		<input type="button" value="2.�귣������" onClick="PopBrandAdminUsingChange('<%= opartner.FPartnerList(i).Fmakerid %>','','');" class="button" <%IF opartner.FPartnerList(i).Fisusing<>"Y" or opartner.FPartnerList(i).Fcurrentusingitemcnt > 0 then %>disabled<%else%>style="color:red;"<%end if%> >
			</div>
			 <%END IF%>
		
			<% if (C_ADMIN_AUTH) then %>
		<% if opartner.FPartnerList(i).IsReqOutProcessBrand() then %>
		    <input type="button" value="����" onClick="brnadOutProc('<%= opartner.FPartnerList(i).Fmakerid %>');" class="button">
		<% elseif opartner.FPartnerList(i).IsReqBrandScmClose() then %>
            <input type="button" value="SCM ����" onClick="scmKillProc('<%= opartner.FPartnerList(i).Fmakerid %>');"  class="button">
		<% end if %>
		<% end if %>
		</td>
	</tr>
	<% next %>
<% end if %>
</table>
 
 <!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table> 
<form name=actfrm method=post action="outbrand_process.asp">
<input type=hidden name="yyyymm" value="<%= yyyy1 %>-<%= mm1 %>">
<input type=hidden name="mode" value="makeoutbrand">
<input type=hidden name="makerid" value="">
</form>
<%
set opartner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->