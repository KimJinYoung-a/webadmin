<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 
' History : ���ʻ����ڸ�
'			2017.04.10 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" -->

<%
''======2010 �߰�=====================
''function optKindSeqCode2Dec(icode)
''    dim ascCode 
''    optKindSeqCode2Dec = icode
''    
''    if (icode>9) then
''        ascCode = ASC(icode)
''        if (ascCode>64) and (ascCode<91) then
''            optKindSeqCode2Dec = CHR(ascCode-55)
''        end if
''    end if
''end function

function optKindSeq2Code(iseq)
    dim ascCode 
    optKindSeq2Code = CStr(iseq)
    
    if (iseq>9) then
        iseq = iseq + 55
        if (iseq>64) and (iseq<91) then
            optKindSeq2Code = CHR(iseq)
        end if
    end if
end function
''======2010 �߰�=====================

dim itemid, optAddType
itemid      = requestCheckVar(request("itemid"),10)
optAddType  = requestCheckVar(request("optAddType"),1)
if optAddType="" then optAddType="N"

dim oitemoption, oOptionMultipleType, oOptionMultiple
dim IsUpchebeasong : IsUpchebeasong = FALSE

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid
if (C_IS_Maker_Upche) then
    oitem.FRectMakerid = session("ssBctid")
end if

if itemid<>"" then
	oitem.GetOneItem
end if

if (oitem.FResultCount<1) then 
    response.write "������ �����ϴ�."
    dbget.close()	:	response.End
end if

IsUpchebeasong = (oitem.FOneItem.Fdeliverytype = "2") or (oitem.FOneItem.Fdeliverytype  = "5") or (oitem.FOneItem.Fdeliverytype  = "9") or (oitem.FOneItem.Fdeliverytype  = "7")
    
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
	
	''���߿ɼ��� �ִ°�� ������..
    if (oitemoption.IsMultipleOption) then
        optAddType = "D"
    end if
end if



set oOptionMultipleType = new CItemOptionMultiple
oOptionMultipleType.FRectItemID = itemid 
if itemid<>"" then
    oOptionMultipleType.GetOptionTypeInfo
end if

set oOptionMultiple = new CitemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" then
    oOptionMultiple.GetOptionMultipleInfo
end if

dim maxcustomoptionno
maxcustomoptionno = 10
for i=0 to oitemoption.FResultCount - 1
    if IsNumeric(oitemoption.FItemlist(i).Fitemoption) then
        if (CInt(oitemoption.FItemlist(i).Fitemoption) < 200) then
            if (CInt(oitemoption.FItemlist(i).Fitemoption) > maxcustomoptionno) then
                maxcustomoptionno = CInt(oitemoption.FItemlist(i).Fitemoption)
            end if
        end if
    end if
next



dim i, j, k, iMaxRows, iMaxCols, found

'' �ִ� ���߿ɼ� 3�� , ���д� 9�� ���� => 19��
iMaxCols = 3
iMaxRows = 30

%>

<script type='text/javascript'>
function SelectAddType(itype){
    var frm = document.frmAddType;
    
    if (((frm.currAddType.value=="N")||(frm.currAddType.value=="S"))){
        location.href="?itemid=" + frm.itemid.value + '&optAddType=' + itype;
        return;
    }
    
    if ((frm.currAddType.value=="D")){
        location.href="?itemid=" + frm.itemid.value + '&optAddType=' + itype;
        return;
    }
    
    var AddType_N   = document.getElementById("optAddType_N");
    var AddType_S   = document.getElementById("optAddType_S");

    if (itype=="N"){
        frm.optAddType[0].checked = true;
        AddType_N.style.display     = "inline";
        AddType_S.style.display     = "none";
    }else if(itype=="S"){
        frm.optAddType[1].checked = true;
        AddType_N.style.display    = "none";
        AddType_S.style.display     = "inline";
    }else if(itype=="D"){
        frm.optAddType[2].checked = true;
    }    
}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : Ÿ��Ʋ �����ϴܿ� searchWrap�� �� ��쿣 bgNone Ŭ���� ���� -->
			<h2>�ɼ��߰�</h2>
			
	    <p class="titDesp tMar05">-	�⺻�ɼ� �߰� : ����, ������� �⺻������ ���ǵ� �ɼ��� �߰� �Ͻ� �� �ֽ��ϴ�.</p>
      <p class="titDesp">- ����ɼ� �߰� : �⺻�ɼ� �̿��� �ɼ��� �߰��Ͻ� �� �ֽ��ϴ�.</p>
      <p class="titDesp">- ���߿ɼ� �߰� : ������ �� ���� �� 2������ �ɼ��� ���� �ؾ� �� ��� ���</p>
      <p class="titDesp">- �ɼǸ� ���� �� �ɼ� ������ [�ɼ� ����] �޴��� �̿��ϼ���.</p>
      <p class="titDesp">- �߰��ݾ��� ������� �ɼ��̸� ����� �Է��Ͻñ� �ٶ��ϴ�. <b>(�ɼǸ� �߰��ݾ��� ���� ������. �ڵ����� ǥ�õ˴ϴ�.)</b></p>
	    </ul> 
	</div>
		<div class="cont">  
<!-- ǥ ��ܹ� ��-->
			<form name="frmAddType">
			<input type="hidden" name="itemid" value="<%= itemid %>">
			<input type="hidden" name="currAddType" value="<%= optAddType %>">
 			<table class="tbType1 writeTb">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<td> 
			    	<span><input type="radio" class="formRadio" name="optAddType" value="N" <%= chkIIF(oitemoption.IsMultipleOption,"disabled","") %> <%= chkIIF(optAddType="N","checked","") %> onClick="SelectAddType('N');"> <a href="javascript:SelectAddType('N');">�⺻�ɼ� �߰�</a></span>
    				<span class="lMar10"><input type="radio"  class="formRadio" name="optAddType" value="S" <%= chkIIF(oitemoption.IsMultipleOption,"disabled","") %> <%= chkIIF(optAddType="S","checked","") %> onClick="SelectAddType('S');"> <a href="javascript:SelectAddType('S');">����ɼ� �߰�</a></span>
    				<span class="lMar10"><input type="radio"  class="formRadio" name="optAddType" value="D" <%= chkIIF(oitemoption.IsMultipleOptionRegAvail,"","disabled") %> <%= chkIIF(optAddType="D","checked","") %> onClick="SelectAddType('D');"> <a href="javascript:SelectAddType('D');">���߿ɼ� �߰�</a></span>
    			</td>
				</tr>
			</tbody>			
			</table>
			</form> 
<script type='text/javascript'>
<!--

var optvalue = (<%= maxcustomoptionno %> + 1); // ����ɼ�(11 - 99)
function saveDoubleOptionAdd(){
//    <!--% if (Not IsUpchebeasong) then %--> //2016.05.19 �ٹ赵 �����ϰ�. ���߿ɼ� ���� ����. ������
//    alert('��ü ����ΰ�츸 ���߿ɼ����� ��ϰ��� �մϴ�.');
//    return;
//    <!--% end if %-->
    
    var frm = document.frmDoubleOption;
    
    if ((fnTrim(frm.optionTypename1.value).length<1)||(fnTrim(frm.optionTypename2.value).length<1)){
        alert('���߿ɼ��� �ɼ� ���и��� �ΰ��� �̻��̾�� �մϴ�.');
        return;
    }
    
    if ((fnTrim(frm.optionTypename1.value)==fnTrim(frm.optionTypename2.value))||(fnTrim(frm.optionTypename2.value)==fnTrim(frm.optionTypename3.value))||(fnTrim(frm.optionTypename1.value)==fnTrim(frm.optionTypename3.value))){
        alert('���߿ɼ��� �ɼ� ���и��� ���� �ٸ��� �����ؾ� �մϴ�.');
        return;
    }

	var oCnt1=0, oCnt2=0, oCnt3=0;
	for(var i=0;i<frm.optionName1.length;i++) {
		
		if(GetByteLength(frm.optionName1[i].value) >30 ){
        	alert("�ɼǸ��� 30byte (�ѱ� 15��, ���� 30��) �̳��� �Է����ּ���");
        	frm.optionName1[i].focus(); 
        	return;
     }
     
		if(fnTrim(frm.optionName1[i].value).length>0) {
			oCnt1++;
		}
	}
	for(var i=0;i<frm.optionName2.length;i++) {
		if(GetByteLength(frm.optionName2[i].value) >30 ){
        	alert("�ɼǸ��� 30byte (�ѱ� 15��, ���� 30��) �̳��� �Է����ּ���");
        	frm.optionName2[i].focus(); 
        	return;
     }
     
		if(fnTrim(frm.optionName2[i].value).length>0) {
			oCnt2++;
		}
	}
	for(var i=0;i<frm.optionName3.length;i++) {
		if(GetByteLength(frm.optionName3[i].value) >30 ){
        	alert("�ɼǸ��� 30byte (�ѱ� 15��, ���� 30��) �̳��� �Է����ּ���");
        	frm.optionName3[i].focus(); 
        	return;
     }
     
		if(fnTrim(frm.optionName3[i].value).length>0) {
			oCnt3++;
		}
	}
	if(oCnt1==0) oCnt1=1;
	if(oCnt2==0) oCnt2=1;
	if(oCnt3==0) oCnt3=1;

	if ((oCnt1*oCnt2*oCnt3)>900){
        alert('�ɼ��� �ʹ� �����ϴ�.\n���� �ɼ��� ����� ���� 900���� ���� �� �����ϴ�.\n\n�ɼ��� ���� �ٿ��ּ���.');
        return;
	}
    
    var ret = confirm('�ɼ��� �߰� �Ͻðڽ��ϱ�?');
	if (ret){
	    frm.submit();
	}
}


function saveItemOptionAdd(){
	var frm = document.frmSelOpt_N;
	var upfrm = document.frmarr;
    
    if (frm.optionTypename.value.length < 1) {
        alert("�ɼ� ���и��� �Է��ϼ���. (ex : ����, ������ ...)");
        frm.optionTypename.focus();
        return;
    }
    
    if (GetByteLength(frm.optionTypename.value)>32){
       alert('�ɼǱ��и��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���'+frm.optionTypename.value); 
       frm.optionTypename.focus();
       return;
     } 
    
    if (frm.addopt.options.length < 1) {
        alert("����� �ɼ��� �����ϴ�.");
        return;
    }

	var ret = confirm('�ɼ��� �߰� �Ͻðڽ��ϱ�?');
	if (ret){
	    upfrm.mode.value = "addoptionCustom";
	    upfrm.optionTypename.value = frm.optionTypename.value;
	    upfrm.arritemoption.value = "";
	    upfrm.arritemoptionname.value = "";
        
        var optCnt = frm.addopt.options.length;
        
        for(var i = 0; i < frm.addopt.options.length; i++) {
            upfrm.arritemoptionname.value += (frm.addopt.options[i].text + "|");

            // ����ɼ��߰�
            if (frm.addopt.options[i].value == "0000") {
                if (optvalue > 299) {
                    alert("�ʹ����� �ɼ��� �߰��ϼ̽��ϴ�..");
                    return;
                }
                frm.addopt.options[i].value = ("000" + optvalue).slice(-4);
                optvalue = optvalue + 1;
            }

            upfrm.arritemoption.value += (frm.addopt.options[i].value + "|");
        }
        upfrm.submit();
    }
}



//�ɼ��������ý� �����ɼ� ����
function searchOption(paramCode1) {

	resetOption1() ;

	if(paramCode1 != '') {
		FrameSearchOption.location.href="/lib/frame_option_select.asp?search_code=" + paramCode1 + "&form_name=frmSelOpt_N&element_name=opt2";
	}
}

//�ɼǸ���Ʈ �ʱ�ȭ
function resetOption1() {
	document.frmSelOpt_N.opt2.length = 1;
	document.frmSelOpt_N.opt2.selectedIndex = 1 ;
}

//���ÿɼ� �ʱ�ȭ
function resetRealOption() {
	opener.document.frmSelOpt_N.addopt.length = 0;
	opener.document.frmSelOpt_N.addopt.selectedIndex = 0 ;
}

function MoveOption(fbox) {
	for(i=0; i<fbox.options.length; i++){
		if(fbox.options[i].selected){
			opener.InsertOption(fbox.options[i].text, fbox.options[i].value)
			fbox.options[i] = null;
			i=i-1;
		}
	}
}

function MoveOptionWithGubun(fbox1,fbox2) {
    var ofrm = document.frmAddType;
    var optionTypename = "";
    
    if (ofrm.optAddType[0].checked){
    
        for(i=0; i<fbox1.options.length; i++){
            if(fbox1.options[i].selected){
                optionTypename = fbox1.options[i].text;
            }
        }
        
    	for(var i=0; i<fbox2.options.length; i++){
    		if(fbox2.options[i].selected){
    		    if (fbox2.options[i].value.length>0){
    			    InsertOptionWithGubun(optionTypename , fbox2.options[i].text, fbox2.options[i].value)
    			    fbox2.options[i] = null;
    			    i=i-1;
    			}
    			
    		}
    	}
    }else if (ofrm.optAddType[1].checked){
    	for (var i=0; i<frmSelOpt_N.etcOpt.length; i++){
	    		  if (GetByteLength(frmSelOpt_N.etcOpt[i].value)>32){
				       alert('�ɼǸ��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���'); 
				       frmSelOpt_N.etcOpt[i].focus();
				       return;
				     } 
	    	}
	    	
        for (var i=0; i<frmSelOpt_N.etcOpt.length; i++){
            InsertOptionEtc(frmSelOpt_N.etcOpt[i].value);
            frmSelOpt_N.etcOpt[i].value = '';
        }
    }
}

function InsertOptionEtc(ioptionText){
    var frm = document.frmSelOpt_N;
    
    var optStr = fnTrim(ioptionText);
    var optcnt = frm.addopt.options.length;
    if (optStr.length>0){
        for (var j=0; j<optcnt; j++){
            if (frm.addopt.options[j].text==optStr) return;
        }
        frm.addopt.options[frm.addopt.options.length] = new Option(optStr, '0000');
    }
}

function InsertOptionWithGubun(optionTypename, ft, fv) {
	var frm = document.frmSelOpt_N;
	var reStr = optionTypename.replace(/\(�ѱ�\)/gi,'');
	
	reStr = reStr.replace(/\(����\)/gi,'');
	reStr = reStr.replace(/\(1-99\)/gi,'');
	reStr = reStr.replace(/����Ŭ��2/gi,'����Ŭ��');
	
	//�̹� �ɼ� ������ ��� ������ �ٲ��� �ʴ´�.
	if (frm.optionTypename.value.length<1){
	    frm.optionTypename.value = reStr;
	}
	
	var optcnt = frm.addopt.options.length;
	for (var j=0; j<optcnt; j++){
        if (frm.addopt.options[j].text==ft) return;
    }
	frm.addopt.options[frm.addopt.options.length] = new Option(ft, fv);
}


// ���õ� �ɼ� ����
function DelSelectOption()
{
	var frm = document.frmSelOpt_N;
	var sidx = frm.addopt.options.selectedIndex;
    var fbox = frm.addopt;
    
	if(sidx<0){
		alert("������ �ɼ��� �������ֽʿ�.");
	}else{
	    for(i=0; i<fbox.options.length; i++){
    		if(fbox.options[i].selected){
    			fbox.options[i] = null;
    			i=i-1;
    		}
    	}
		
		if (fbox.options.length<1){
		    frm.optionTypename.value = '';
		}
	}
}
//-->
</script>

<% if optAddType<>"D" then %>
<!-- ���Ͽɼ� -->
 
        <form name="frmSelOpt_N">
        <table class="tbType1 listTb tMar10">
        <tr>
            <td width="460" >
                <div id="optAddType_N" <%= chkIIF(optAddType="N" ,"style='display:inline'","style='display:none'") %>>
                 <table class="tbType1 listTb">
                	<tr>
                		<th><div>�ɼ� ����</div></th>
                		<th><div>�ɼ� ��</div></th>
                	</tr>
                	<tr>
                		<td>
                		  <select class="formSlt" multiple="multiple" name="opt1" size="20" style='width:210px;height:400px' onchange="javascript:searchOption(this.options[this.selectedIndex].value);" >
                		  <option value="">-----------------------</option>
                		  </select>
                		</td>
                		<td>
                		  <select class="formSlt" multiple="multiple" name="opt2" size="20" style='width:210px;height:400px'>
                		  <option value="">-----------------------</option>
                		  </select> 
                		</td>
                	</tr>
                	<!--
                	<tr bgcolor="#FFFFFF">
                		<td colspan="4" align="center">
                			<input type="button" value="���ÿɼ��߰�" onclick="MoveOptionWithGubun(document.itemopt.elements['opt1'],document.itemopt.elements['opt2'])">
                			<input type="button" value=" �� �� " onclick="self.close()">
                		</td>
                	</tr>
                	-->
                    
                </table>
                </div>
                <div id="optAddType_S" <%= chkIIF(optAddType="S" ,"style='display:inline'","style='display:none'") %>>
                <table  class="tbType1 listTb">
                <% for i=0 to 9 %>
                <tr>
                    <th><div> �ɼǸ� <%= i+1 %> </div></th>
                    <td><input type="text"class="formTxt" name="etcOpt" size="20" maxlength="20"></td>
                </tr>
                <% next %>
                </table>
                </div>
            </td>
            <td width="50">
                <input type="button" class="btn cRd1" value=">> �߰�" onclick="MoveOptionWithGubun(document.frmSelOpt_N.elements['opt1'],document.frmSelOpt_N.elements['opt2'])">
                <br><br>
                <input type="button" class="btn" value="<< ����" onclick="DelSelectOption();">
            </td>
            <td width="300">
                <table  class="tbType1 listTb"> 
                <tr>
                    <th><div>�߰� �� �ɼ�</div></th>
                </tr> 
                <tr>
                    <td>�ɼǱ��и�  
                    <% if (oitemOption.FResultCount<1) then %>
                    <input type="text" class="formTxt" name="optionTypename" value="" size="19" maxlength="20">
                    <% else %>
                    <input type="text" class="formTxt" name="optionTypename" value="<%= oitemOption.FItemList(0).FoptionTypeName %>" size="19" maxlength="20">
                    <% end if %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <select multiple="multiple" name="addopt" size="8" style='width:210px;height:150px'>
        		        </select>
        		    </td>
        		</tr>
        		<tr>
                    <th><div>���� ��ϵ� �ɼ�</div></th>
                </tr>
                <tr >
                    <td>
                        <select multiple name="oldopt" size="8" style='width:210px;height:150px;background-color:#CCCCCC' >
                        <% for i=0 to oitemoption.FResultCount-1 %>
                        <option value="<%= oitemoption.FItemList(i).Fitemoption %>"><%= oitemoption.FItemList(i).FoptionName %>
                        <% next %>
        		        </select>
        		    </td>
        		</tr>
        		</table>
            </td>
           </tr>
         </table>
         </form>
    </td>
</tr>
</table>
<iframe name="FrameSearchOption" src="/lib/frame_option_select.asp?form_name=frmSelOpt_N&element_name=opt1" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>

<% Else %>
<!-- ���߿ɼ� -->
    <% if (oitemoption.IsMultipleOption) or (oitemoption.FResultCount<1) then %>
   	<ul  class="txtList">
   	 <li>�ɼǱ��и� : ����, ������ ��.. �Է�</li>
     <li>�ɼǸ� : ����, ���, �Ķ� ��.. �Է�</li>
    </ul>     	
   	<form name="frmDoubleOption" method="post" action="do_adminitemoptionedit.asp">
   	<input type="hidden" name="mode" value="addDoubleOption">
   	<input type="hidden" name="itemid" value="<%= itemid %>">
    <table class="tbType1 listTb tMar10"> 
    	<thead>
    	<tr>
        <th><div>�ɼǱ��и�
        <% for j=0 to iMaxCols-1 %></div></th>
        <th><div>
            <% found = FALSE %>
            <% for k=0 to oOptionMultipleType.FResultCount-1 %> 
                <% if (oOptionMultipleType.FItemList(k).FTypeSeq=(j+1)) then %>
                <input type="text" name="optionTypename<%= j+1 %>" class="formTxt readonly" ReadOnly value="<%= oOptionMultipleType.FItemList(k).FOptionTypeName %>" size="18" maxlength="20">
                <% found = TRUE %>
                <% end if %>
            <% next %>
            
            <% if Not found then %>
            <input type="text" class="formTxt" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20">
            <% end if %>
        </div></th>
        <% Next %>
        <th><div>(��Ͽ���)<br/>����</div></th>
        <th><div>(��Ͽ���)<br/>������</div></th>
    </tr>
    </thead>
    <tbody>    
    <% for i=0 to iMaxRows-1 %>
    <tr>
        <td>�ɼǸ� <%= i+1 %></td>
        <% for j=0 to iMaxCols-1 %>
        <td>
            <% found = FALSE %>
            <% for k=0 to oOptionMultiple.FResultCount-1 %>
                <% if (oOptionMultiple.FItemList(k).FTypeSeq=(j+1)) and (CStr(oOptionMultiple.FItemList(k).FKindSeq)=optKindSeq2Code(i+1)) then %>
                    <input type="hidden" name="itemoption<%= j+1 %>" value="<%= oOptionMultiple.FItemList(k).FTypeSeq %><%= oOptionMultiple.FItemList(k).FKindSeq %>">
                    <input type="text" name="optionName<%= j+1 %>" class="formTxt readonly" ReadOnly value="<%= oOptionMultiple.FItemList(k).FoptionKindName %>" size="16" maxlength="20">
                    <% found = TRUE %>
                <% end if %>
            <% next %>
            
            <% if Not found then %>
            <input type="hidden" name="itemoption<%= j+1 %>" value="">
            <input type="text" class="formTxt" name="optionName<%= j+1 %>" size="16" maxlength="20">
            <% end if %>
        </td>
        <% next %>
        <td>
            <% if i=0 then %>
            ����
            <% elseif i=1 then %>
            �Ķ�
            <% elseif i=2 then %>
            ���
            <% elseif i=3 then %>
            ������
            <% end if %>
        </td>
        <td>
            <% if i=0 then %>
            XL
            <% elseif i=1 then %>
            L
            <% elseif i=2 then %>
            S
            <% end if %>
        </td>
    </tr>
    
    <% next %>
  </tbody>
    <!-- �� �߰� ����. iMaxRows ����
    <tr bgcolor="#FFFFFF">
        <td align="center" ><a href="javascript:AddRows();"><img src="/images/icon_plus.gif" width="16" alt="�� �߰�" border="0"></a></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    -->
    </table> 
    </form>
    
    <% else %>
    <!-- ���� ���� �ɼ��� ��� �Ұ� : ���� �� ����.-->
    <div class="titDesp tMar05 cRd1"> ���� �ɼ� �߰� �Ұ�. </div>
    <div class="titDesp tMar05">���߿ɼ��� ����Ͻ÷��� ���� ���� �ɼ����� ��ϵ� �ɼ��� ���� ���� �� �����մϴ�.</div> 
    <% end if %>
<% end if %> 
 			<div class="tPad15 ct"> 
          <% if optAddType<>"D" then %>
          <input type="button" value="�߰��� �ɼ� ����" name="btnoptsave" onclick="saveItemOptionAdd();" class="btn3 btnRd"/>
          <% else %>
          <input type="button" value="�߰��� �ɼ� ����" name="btnoptsave" onclick="saveDoubleOptionAdd();" class="btn3 btnRd"/>
          <% end if %> 
			</div>  
<form name=frmarr method=post action="do_adminitemoptionedit.asp">
<input type=hidden name=mode value="">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=optionTypename value="">
<input type=hidden name=arritemoption value="">
<input type=hidden name=arritemoptionname value="">
</form>
<%
set oitem = Nothing
set oOptionMultipleType = Nothing
set oOptionMultiple = Nothing
set oitemoption = Nothing
%>
 		</div>
	</div>
</div>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->