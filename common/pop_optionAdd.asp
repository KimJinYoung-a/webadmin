<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �ɼǵ��
' History : ���� ������ ��
'			2017.04.10 �ѿ�� ����(����ó��)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

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
        if (CInt(oitemoption.FItemlist(i).Fitemoption) < 9998) then
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

<script language='javascript'>
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
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td valign="top" bgcolor="F4F4F4">
	        	<b>�ɼ��߰�</b><br>

	        	<br>- �⺻�ɼ� �߰� : ����, ������� �⺻������ ���ǵ� �ɼ��� �߰� �Ͻ� �� �ֽ��ϴ�.
                <br>- ����ɼ� �߰� : �⺻�ɼ� �̿��� �ɼ��� �߰��Ͻ� �� �ֽ��ϴ�.
                <br>- ���߿ɼ� �߰� : ������ �� ���� �� 2������ �ɼ��� ���� �ؾ� �� ��� ���
                <br>- �ɼǸ� ���� �� �ɼ� ������ [�ɼ� ����] �޴��� �̿��ϼ���.
                <br>- �߰��ݾ��� ������� �ɼ��̸� ����� �Է��Ͻñ� �ٶ��ϴ�. <b>(�ɼǸ� �߰��ݾ��� ���� ������. �ڵ����� ǥ�õ˴ϴ�.)</b>
	        </td>
	</tr>
	</form>
</table>
<p>
<!-- ǥ ��ܹ� ��-->

<table border="0" cellspacing="1" cellpadding="2" width="100%" class="a" bgcolor="#BABABA">
<form name="frmAddType">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="currAddType" value="<%= optAddType %>">
<tr>
    <td colspan="2" bgcolor="#FFFFFF" height="30">
    <input type="radio" name="optAddType" value="N" <%= chkIIF(oitemoption.IsMultipleOption,"disabled","") %> <%= chkIIF(optAddType="N","checked","") %> onClick="SelectAddType('N');"><a href="javascript:SelectAddType('N');">�⺻�ɼ� �߰�</a>
    &nbsp;
    <input type="radio" name="optAddType" value="S" <%= chkIIF(oitemoption.IsMultipleOption,"disabled","") %> <%= chkIIF(optAddType="S","checked","") %> onClick="SelectAddType('S');"><a href="javascript:SelectAddType('S');">����ɼ� �߰�</a>
    &nbsp;
    <input type="radio" name="optAddType" value="D" <%= chkIIF(oitemoption.IsMultipleOptionRegAvail,"","disabled") %> <%= chkIIF(optAddType="D","checked","") %> onClick="SelectAddType('D');"><a href="javascript:SelectAddType('D');">���߿ɼ� �߰�</a>
    </td>
</tr>
</form>
</table>

<p>


<script language="JavaScript">
<!--

var optvalue = (<%= maxcustomoptionno %> + 1); // ����ɼ�(0011 - 9999)
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
		if(fnTrim(frm.optionName1[i].value).length>0) {
			oCnt1++;
		}
	}
	for(var i=0;i<frm.optionName2.length;i++) {
		if(fnTrim(frm.optionName2[i].value).length>0) {
			oCnt2++;
		}
	}
	for(var i=0;i<frm.optionName3.length;i++) {
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
                if (optvalue > 9999) {
                    alert("�ʹ����� �ɼ��� �߰��ϼ̽��ϴ�.");
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
<table border="0" cellspacing="1" cellpadding="2" width="100%" class="a" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td>
        <form name="frmSelOpt_N">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#FFFFFF">
        <tr>
            <td width="460" >
                <div id="optAddType_N" <%= chkIIF(optAddType="N" ,"style='display:inline'","style='display:none'") %>>
                <table width="440" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
                	<tr height="30" bgcolor="#DDDDFF" align="center">
                		<td>�ɼ� ����</td>
                		<td>�ɼ� ��</td>
                	</tr>
                	<tr bgcolor="#FFFFFF" align="center">
                		<td>
                		  <select name="opt1" size="20" style='width:210;' onchange="javascript:searchOption(this.options[this.selectedIndex].value);" >
                		  <option value="">-----------------------</option>
                		  </select>
                		</td>
                		<td>
                		  <select multiple name="opt2" size="20" style='width:210;'>
                		  <option value="">-----------------------</option>
                		  </select>&nbsp;
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
                <table width="440" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d" >
                <% for i=0 to 9 %>
                <tr bgcolor="#FFFFFF" align="center">
                    <td>�ɼǸ� <%= i+1 %> </td>
                    <td align="center"><input type="text" name="etcOpt" size="20" maxlength="40"></td>
                </tr>
                <% next %>
                </table>
                </div>
            </td>
            <td width="50">
                <input type="button" value=">> �߰�" onclick="MoveOptionWithGubun(document.frmSelOpt_N.elements['opt1'],document.frmSelOpt_N.elements['opt2'])">
                <br><br>
                <input type="button" value="<< ����" onclick="DelSelectOption();">
            </td>
            <td width="300">
                <table width="220" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
                <tr height="30" bgcolor="#DDDDFF" align="center">
                    <td>�߰� �� �ɼ�</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td>�ɼǱ��и�  
                    <% if (oitemOption.FResultCount<1) then %>
                    <input type="text" name="optionTypename" value="" size="19" maxlength="20">
                    <% else %>
                    <input type="text" name="optionTypename" value="<%= oitemOption.FItemList(0).FoptionTypeName %>" size="19" maxlength="20">
                    <% end if %>
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center">
                        <select multiple name="addopt" size="8" style='width:210;'>
        		        </select>
        		    </td>
        		</tr>
        		<tr height="20" bgcolor="#DDDDFF" align="center">
                    <td>���� ��ϵ� �ɼ�</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center">
                        <select multiple name="oldopt" size="8" style='width:210;background-color:#CCCCCC' >
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
    <table border="0" cellspacing="1" cellpadding="2" width="100%" class="a" bgcolor="#BABABA">
    <form name="frmDoubleOption" method="post" action="do_adminitemoptionedit.asp">
    <input type="hidden" name="mode" value="addDoubleOption">
    <input type="hidden" name="itemid" value="<%= itemid %>">
    <tr bgcolor="#FFFFFF">
        <td> 
            &nbsp;&nbsp;- �ɼǱ��и� : ����, ������ ��.. �Է�
            <br>
            &nbsp;&nbsp;- �ɼǸ� : ����, ���, �Ķ� ��.. �Է�
        </td>
    </tr>
    <tr bgcolor="#FFFFFF">
        <td>
            <table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
            <tr align="center"  bgcolor="#DDDDFF">
                <td width="100">�ɼǱ��и�</td>
                <% for j=0 to iMaxCols-1 %>
                <td> 
                    <% found = FALSE %>
                    <% for k=0 to oOptionMultipleType.FResultCount-1 %> 
                        <% if (oOptionMultipleType.FItemList(k).FTypeSeq=(j+1)) then %>
                        <input type="text" name="optionTypename<%= j+1 %>" class="text_ro" ReadOnly value="<%= oOptionMultipleType.FItemList(k).FOptionTypeName %>" size="18" maxlength="20">
                        <% found = TRUE %>
                        <% end if %>
                    <% next %>
                    
                    <% if Not found then %>
                    <input type="text" name="optionTypename<%= j+1 %>" value="" size="18" maxlength="20">
                    <% end if %>
                </td>
                <% Next %>
                <td width="80">(��Ͽ���)<br>����</td>
                <td width="80">(��Ͽ���)<br>������</td>
            </tr>
            <tr height="2" bgcolor="#FFFFFF">
                <td colspan="6"></td>
            </tr>
            <% for i=0 to iMaxRows-1 %>
            <tr align="center"  bgcolor="#FFFFFF">
                <td>�ɼǸ� <%= i+1 %></td>
                <% for j=0 to iMaxCols-1 %>
                <td>
                    <% found = FALSE %>
                    <% for k=0 to oOptionMultiple.FResultCount-1 %>
                        <% if (oOptionMultiple.FItemList(k).FTypeSeq=(j+1)) and (CStr(oOptionMultiple.FItemList(k).FKindSeq)=optKindSeq2Code(i+1)) then %>
                            <input type="hidden" name="itemoption<%= j+1 %>" value="<%= oOptionMultiple.FItemList(k).FTypeSeq %><%= oOptionMultiple.FItemList(k).FKindSeq %>">
                            <input type="text" name="optionName<%= j+1 %>" class="text_ro" ReadOnly value="<%= oOptionMultiple.FItemList(k).FoptionKindName %>" size="16" maxlength="20">
                            <% found = TRUE %>
                        <% end if %>
                    <% next %>
                    
                    <% if Not found then %>
                    <input type="hidden" name="itemoption<%= j+1 %>" value="">
                    <input type="text" name="optionName<%= j+1 %>" size="16" maxlength="20">
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
            
        </td>
    </tr>
    </form>
    </table>
    <% else %>
    <!-- ���� ���� �ɼ��� ��� �Ұ� : ���� �� ����.-->
    <table border="0" cellspacing="1" cellpadding="2" width="100%" class="a" bgcolor="#BABABA">
    <tr bgcolor="#FFFFFF" height="50">
        <td align="center">
            ���� �ɼ� �߰� �Ұ�. 
            <br> ���߿ɼ��� ����Ͻ÷��� ���� ���� �ɼ����� ��ϵ� �ɼ��� ���� ���� �� �����մϴ�.
        </td>
    </tr>
    </table>
    <% end if %>
<% end if %>



<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
          <% if optAddType<>"D" then %>
          <input type="button" value="�߰��� �ɼ� ����" name="btnoptsave" onclick="saveItemOptionAdd();">
          <% else %>
          <input type="button" value="�߰��� �ɼ� ����" name="btnoptsave" onclick="saveDoubleOptionAdd();">
          <% end if %>
          <input type="button" value=" �� �� " onclick="window.close()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->



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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->