<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/waititemcls_2008.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<%
dim mode
dim itemid, itemoption
dim oitem, oitemoption, oOptionMultipleType, oOptionMultiple

itemid = requestCheckvar(request("itemid"),16)
if itemid="" then itemid=0
mode= requestCheckvar(request("mode"),32)
itemoption= requestCheckvar(request("itemoption"),4)

dim sqlStr
dim ErrStr

set oitem = new CWaitItem
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

set oitemoption = new CWaitItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

set oOptionMultipleType = new CWaitItemOptionMultiple
oOptionMultipleType.FRectItemID = itemid 
if itemid<>"" then
    oOptionMultipleType.GetOptionTypeInfo
end if

set oOptionMultiple = new CWaititemOptionMultiple
oOptionMultiple.FRectItemID = itemid
if itemid<>"" then
    oOptionMultiple.GetOptionMultipleInfo
end if
        

dim i, j, k, TrFlag, pp
TrFlag = false
pp=0

dim maxcustomoptionno
maxcustomoptionno = 11
for i=0 to oitemoption.FResultCount - 1
    if IsNumeric(oitemoption.FItemlist(i).Fitemoption) then
        if (CInt(oitemoption.FItemlist(i).Fitemoption) < 100) then
            if (CInt(oitemoption.FItemlist(i).Fitemoption) > maxcustomoptionno) then
                maxcustomoptionno = CInt(oitemoption.FItemlist(i).Fitemoption)
            end if
        end if
    end if
next

dim ItemDefaultMargin
ItemDefaultMargin = 100-CLng(oitem.FOneItem.FBuycash/oitem.FOneItem.Fsellcash*100*100)/100


%>

<script type='text/javascript'>
var VItemDefaultMargin = <%= ItemDefaultMargin %>;

function EditOptionInfo(){
    var frm = document.frmEdit;
    var optAddpriceExists = false;
    
    if (frm.mode.value=="editOptionMultiple"){
        //���߿ɼ�
        if (!frm.optionTypename.length){
            if (frm.optionTypename.value.length<1){
                alert('�ɼ� ���и��� �Է��ϼ���.');
                frm.optionTypename.focus();
                return;
            }
            
            if (GetByteLength(frm.optionTypename.value)>32){
                alert('�ɼǱ��и��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���'); 
                frm.optionTypename.focus();
                return;
            } 
        }else{
            for (var i=0;i<frm.optionTypename.length;i++){
                if (frm.optionTypename[i].value.length<1){
                    alert('�ɼ� ���и��� �Է��ϼ���.');
                    frm.optionTypename[i].focus();
                    return;
                }
                
                if (GetByteLength(frm.optionTypename[i].value)>32){
                alert('�ɼǱ��и��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���'); 
                frm.optionTypename[i].focus();
                return;
            } 
            
                //�ɼǱ��и��� �ߺ��Ǵ��� üũ.
                for (var j=0;j<frm.optionTypename.length;j++){
                    if ((i!=j)&&(fnTrim(frm.optionTypename[i].value)==fnTrim(frm.optionTypename[j].value))){
                        alert('�ɼ� ���и��� �ߺ��Ͽ� ����� �� �����ϴ�. - [' + frm.optionTypename[j].value + ']');
                        frm.optionTypename[j].focus();
                        return;
                    }
                }
            }
        }
        
        if (!frm.optionName.length){
            if (frm.optionName.value.length<1){
                alert('�ɼǸ��� �Է��ϼ���.');
                frm.optionName.focus();
                return;
            }
            
             if(GetByteLength(frm.optionName.value) >30 ){
			        	alert("�ɼǸ��� 30byte (�ѱ� 15��, ���� 30��) �̳��� �Է����ּ���");
			        	frm.optionName.focus(); 
			        	return;
			        } 
        }else{
            for (var i=0;i<frm.optionName.length;i++){
                if (frm.optionName[i].value.length<1){
                    alert('�ɼǸ��� �Է��ϼ���.');
                    frm.optionName[i].focus();
                    return;
                }
                
                if(GetByteLength(frm.optionName[i].value) >30 ){
				        	alert("�ɼǸ��� 30byte (�ѱ� 15��, ���� 30��) �̳��� �Է����ּ���");
				        	frm.optionName[i].focus(); 
				        	return;
				        }
			        
                //�ɼǸ��� �ߺ��Ǵ��� üũ.(���߿ɼ��϶� �ɼǻ󼼸� �ߺ������ϹǷ� ���� : (frm.TypeSeq[i].value==frm.TypeSeq[j].value) �����߰�)
                for (var j=0;j<frm.optionName.length;j++){
                    if ((i!=j)&&(fnTrim(frm.optionName[i].value)==fnTrim(frm.optionName[j].value))&&(frm.TypeSeq[i].value==frm.TypeSeq[j].value)){
                        alert('�ɼǸ��� �ߺ��Ͽ� ����� �� �����ϴ�. - [' + frm.optionName[j].value + ']');
                        frm.optionName[j].focus();
                        return;
                    }
                }
                
            }
        }
        
        //�߰��ݾ�
        if (!frm.optaddprice.length){
            if (frm.optaddprice.value.length<1){
                alert('�߰��ݾ��� �Է��ϼ���.');
                frm.optaddprice.focus();
                return;
            }

             if(frm.optaddprice.value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }

            
            if (!IsDigit(frm.optaddprice.value)){
                alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                frm.optaddprice.focus();
                return;
            }
            
            if ((frm.optaddbuyprice.value*1)>(frm.optaddprice.value*1)) {
                alert('���ް��� ���԰� ���� Ŭ �� �����ϴ�.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if ((frm.optaddprice.value*1>0) && (frm.optaddbuyprice.value*1!=parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100))) {
                if (!confirm('�ɼ� �߰� �ݾ׿� ���� ���� �ݾ��� ��ǰ �⺻ ���� (<%= ItemDefaultMargin %>) ���޾�(' + parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100) + '��) �� ��ġ ���� �ʽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                    frm.optaddbuyprice.focus();
                    return;
                }
            }
            optAddpriceExists = (optAddpriceExists||(frm.optaddprice.value*1>0));
        }else{
            for (var i=0;i<frm.optaddprice.length;i++){
                if (frm.optaddprice[i].value.length<1){
                    alert('�߰��ݾ��� �Է��ϼ���.');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if(frm.optaddbuyprice[i].value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }


                if (!IsDigit(frm.optaddprice[i].value)){
                    alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if ((frm.optaddbuyprice[i].value*1)>(frm.optaddprice[i].value*1)) {
                    alert('���ް��� ���԰� ���� Ŭ �� �����ϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if ((frm.optaddprice[i].value*1>0) && (frm.optaddbuyprice[i].value*1!=parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100))) {
                    if (!confirm('�ɼ� �߰� �ݾ׿� ���� ���� �ݾ��� ��ǰ �⺻ ���� (<%= ItemDefaultMargin %>) ���޾�(' + parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100) + '��) �� ��ġ ���� �ʽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                }
                
                optAddpriceExists = (optAddpriceExists||(frm.optaddprice[i].value*1>0));
            }
        }
        
        //�߰��ݾ�-���԰�
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('�߰��ݾ� ���԰��� �Է��ϼ���.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddbuyprice.value)){
                alert('�߰��ݾ� ���԰��� ���ڸ� �����մϴ�.');
                frm.optaddbuyprice.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('�߰��ݾ� ���԰��� �Է��ϼ���.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddbuyprice[i].value)){
                    alert('�߰��ݾ� ���԰��� ���ڸ� �����մϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
            }
        }

        //�߰��ݾ� ���� �⺻�ɼ� ���翩�� Ȯ��
        if (!frm.optaddbuyprice.length){
            if (frm.optaddprice.value>0){
                alert('�ɼǱ��� ���� �⺻�ɼ��� �ʿ��մϴ�.\n�߰��ݾ��� ����(0��) �⺻ �ɼ��� �߰����ּ���.');
                return;
            }
        }else{
            var chkPreTseq, chkBsOpt = false;
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if(chkPreTseq != frm.TypeSeq[i].value) chkBsOpt = false;
                chkPreTseq = frm.TypeSeq[i].value
                if (frm.optaddprice[i].value==0){
                    chkBsOpt = true;
                }
            }

            if(!chkBsOpt) {
                alert('�ɼǱ��� ���� �⺻�ɼ��� �ʿ��մϴ�.\n�߰��ݾ��� ����(0��) �⺻ �ɼ��� �߰����ּ���.');
                return;
            }
        }
    }else{
        //���Ͽɼ�
        if (frm.optionTypename.value.length<1){
            alert('�ɼ� ���и��� �Է��ϼ���.');
            frm.optionTypename.focus();
            return;
        }
        
          if (GetByteLength(frm.optionTypename.value)>32){
                alert('�ɼǱ��и��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���');
                frm.optionTypename.focus();
                return;
            } 
            
        if (!frm.optionName.length){
            if (frm.optionName.value.length<1){
                alert('�ɼǸ��� �Է��ϼ���.');
                frm.optionName.focus();
                return;
            }
            
            if(GetByteLength(frm.optionName.value) >32 ){
				        	alert("�ɼǸ��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���");
				        	frm.optionName.focus(); 
				        	return;
				        }
        }else{
            for (var i=0;i<frm.optionName.length;i++){
                if (frm.optionName[i].value.length<1){
                    alert('�ɼǸ��� �Է��ϼ���.');
                    frm.optionName[i].focus();
                    return;
                }
                
               if(GetByteLength(frm.optionName[i].value) >32 ){
				        	alert("�ɼǸ��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���");
				        	frm.optionName[i].focus(); 
				        	return;
				        }
				        
                //�ɼǸ��� �ߺ��Ǵ��� üũ.
                for (var j=0;j<frm.optionName.length;j++){
                    if ((i!=j)&&(frm.optionName[i].value==frm.optionName[j].value)){
                        alert('�ɼǸ��� �ߺ��Ͽ� ����� �� �����ϴ�. - [' + frm.optionName[j].value + ']');
                        frm.optionName[j].focus();
                        return;
                    }
                }
                
            }
            
            
        }
        
        //�߰��ݾ�
        if (!frm.optaddprice.length){
            if (frm.optaddprice.value.length<1){
                alert('�߰��ݾ��� �Է��ϼ���.');
                frm.optaddprice.focus();
                return;
            }
            
            if(frm.optaddprice.value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }
                  

            if (!IsDigit(frm.optaddprice.value)){
                alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                frm.optaddprice.focus();
                return;
            }
            
            if ((frm.optaddbuyprice.value*1)>(frm.optaddprice.value*1)) {
                alert('���ް��� ���԰� ���� Ŭ �� �����ϴ�.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if ((frm.optaddprice.value*1>0) && (frm.optaddbuyprice.value*1!=parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100))) {
                if (!confirm('�ɼ� �߰� �ݾ׿� ���� ���� �ݾ��� ��ǰ �⺻ ���� (<%= ItemDefaultMargin %>) ���޾�(' + parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100) + '��) �� ��ġ ���� �ʽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                    frm.optaddbuyprice.focus();
                    return;
                }
            }
            
            optAddpriceExists = (optAddpriceExists||(frm.optaddprice.value*1>0));
        }else{
            for (var i=0;i<frm.optaddprice.length;i++){
                if (frm.optaddprice[i].value.length<1){
                    alert('�߰��ݾ��� �Է��ϼ���.');
                    frm.optaddprice[i].focus();
                    return;
                }

                 if(frm.optaddprice[i].value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                } 
                
                if (!IsDigit(frm.optaddprice[i].value)){
                    alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                    frm.optaddprice[i].focus();
                    return;
                }
                
                if ((frm.optaddbuyprice[i].value*1)>(frm.optaddprice[i].value*1)) {
                    alert('���ް��� ���԰� ���� Ŭ �� �����ϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if ((frm.optaddprice[i].value*1>0) && (frm.optaddbuyprice[i].value*1!=parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100))) {
                    if (!confirm('�ɼ� �߰� �ݾ׿� ���� ���� �ݾ��� ��ǰ �⺻ ���� (<%= ItemDefaultMargin %>) ���޾�(' + parseInt(frm.optaddprice[i].value*1*(100-VItemDefaultMargin)/100) + '��) �� ��ġ ���� �ʽ��ϴ�. ��� �Ͻðڽ��ϱ�?')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                }
                
                optAddpriceExists = (optAddpriceExists||(frm.optaddprice[i].value*1>0));
            }
        }
        
        //�߰��ݾ�-���԰�
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('�߰��ݾ� ���԰��� �Է��ϼ���.');
                frm.optaddbuyprice.focus();
                return;
            }
            
            if (!IsDigit(frm.optaddbuyprice.value)){
                alert('�߰��ݾ� ���԰��� ���ڸ� �����մϴ�.');
                frm.optaddbuyprice.focus();
                return;
            }
        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('�߰��ݾ� ���԰��� �Է��ϼ���.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
                
                if (!IsDigit(frm.optaddbuyprice[i].value)){
                    alert('�߰��ݾ� ���԰��� ���ڸ� �����մϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                }
            }
        }

        //�߰��ݾ� ���� �⺻�ɼ� ���翩�� Ȯ��
        if (!frm.optaddbuyprice.length){
            if (frm.optaddprice.value>0){
                alert('�⺻�ɼ��� �ʿ��մϴ�.\n�߰��ݾ��� ����(0��) �⺻ �ɼ��� �߰����ּ���.');
                return;
            }
        }else{
            var chkBsOpt = false;
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddprice[i].value==0){
                    chkBsOpt = true;
                }
            }

            if(!chkBsOpt) {
                alert('�⺻�ɼ��� �ʿ��մϴ�.\n�߰��ݾ��� ����(0��) �⺻ �ɼ��� �߰����ּ���.');
                return;
            }
        }
    }
    
    <% if (oitem.FOneItem.FMwDiv<>"U") then %>
    //�ٹ���� �ɼ� �߰��ݾ� ���Ұ� �ϰ� 20130228
    if (optAddpriceExists){
        alert('�ٹ����� ����� ��� �ɼ� �߰��ݾ��� ����� �� �����ϴ�.');
        return;
    }
    <% end if %>
    if (optAddpriceExists){
    	var isOversea = "<%=oitem.FOneItem.FdeliverOverseas%>";
    	if (isOversea=="Y"){
    		   alert('�ؿܹ���� �ϴ� ��� �ɼ� �߰��ݾ��� ����� �� �����ϴ�.');
        return;
    	}
    }
  
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}


function SaveOption(){
	var frm;
	var upfrm = document.frmarr;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
					upfrm.itemid.value = upfrm.itemid.value + "|" + frm.itemid.value;
					upfrm.itemoption.value = upfrm.itemoption.value + "|" + frm.itemoption.value;
					if (frm.isusing[0].checked==true){
						upfrm.isusing.value = upfrm.isusing.value + "|" + "Y";
					}else{
						upfrm.isusing.value = upfrm.isusing.value + "|" + "N";
					}
			}
		}
		upfrm.mode.value = "modiitemoptionarr";
		upfrm.submit();
	}
}

function DelItemOption(itemid,itemoption){
    var frm = document.frmOption;
    
	if (confirm('��ǰ ������ ������� �ʴ��� �������� ���ñ� �ٶ��ϴ�. \n\n���� ���� �Ͻðڽ��ϱ�?')){
		frm.mode.value = "deleteoption";
		frm.itemid.value = itemid;
		frm.itemoption.value = itemoption;
		frm.submit();
	}
}


function DelItemOptionMultiple(itemid,typeseq,kindseq){
    var frm = document.frmOption;
    
    if (confirm('��ǰ ������ ������� �ʴ��� �������� ���ñ� �ٶ��ϴ�. \n\n���� ���� �Ͻðڽ��ϱ�?')){
		frm.mode.value = "deleteMultipleOption";
		frm.itemid.value = itemid;
		frm.typeseq.value = typeseq;
		frm.kindseq.value = kindseq;
		frm.submit();
	}
}

function AutoCalcuBuyPrice(comp,j){
    var frm = document.frmEdit;
    
    if (!frm.optaddbuyprice.length){
        frm.optaddbuyprice.value = parseInt(frm.optaddprice.value*1*(100-VItemDefaultMargin)/100);
    }else{
        frm.optaddbuyprice[j].value = parseInt(frm.optaddprice[j].value*1*(100-VItemDefaultMargin)/100);
    }

}
// ============================================================================


function AddOptionPop(iitemid){
    var popwin = window.open('pop_upchewaititemoptionAdd.asp?itemid=' + iitemid,'pop_upchewaititemoptionAdd','width=900,height=900,scrollbars=yes,resizable=yes');
    popwin.focus();
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
			<h2>�ɼǼ���</h2>
			<p class="titDesp tMar05">- �ɼ��� �߰� �Ǵ� �����Ҽ� �ֽ��ϴ�.</p>
			<p class="titDesp">- �Ǹ�/�԰�/���� ������ �ִ� �ɼ��� ������ �Ұ����մϴ�.(������ ���� �����ϼ���)</p>
			</div>
		<div class="cont"> 
		 <table class="tbType1 writeTb tMar10">
				<colgroup>
					<col width="15%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>��ǰ�ڵ�</div></th>
					<td><%= itemid %></td>
					<th width="240">�ɼ� ���� �̸�����</th>
				</tr>
				<tr>
					<th><div>��ǰ��</div></th>
					<td><%= oitem.FOneItem.Fitemname %></td>
					<td  rowspan="2" align="center" style="background-color:#FFF2E6">
					<%= getOptionBoxHTML_FrontType(itemid) %>
					</td>
				</tr>
				<tr>
					<th><div>�귣��</div></th>
					<td><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td> 
				</tr>
				<tbody>
			</table> 
			
			<form name="frmEdit" method="post" action="do_waititemoptionedit.asp">
			<input type="hidden" name="itemid" value="<%= itemid %>">
			<% if (oitemoption.IsMultipleOption) then %>
			<input type="hidden" name="mode" value="editOptionMultiple">
			<% else %>
			<input type="hidden" name="mode" value="editOption">
			<% end if %>
			<div class="overHidden">
			 <div class="ftLt"><h3 class="tMar20 bltNo">��ϵ� �ɼ� ����Ʈ</h3></div>
 			 <div class="ftRt tMar10"><input type="button" class="btn3 btnIntb cBl1" value="�ɼ��߰� +" onClick="AddOptionPop('<%= itemid %>');"></div>
 			</div>
 			 <% if oitemoption.FResultCount<1 then %>
 			 	<ul class="txtList">
					<li class="cBl1">��ϵ� �ɼ��� �����ϴ�.</li>
				</ul>
 			 	<%else%>
 			 	  <% if (oitemoption.IsMultipleOption) then %><!-- ���߿ɼ� --> 
					    <table class="tbType1 listTb tMar10">
							<thead>
							<tr>
								<th><div>����</div></th>
								<th><div>�ɼǱ��и�</div></th>
								<th><div>�ɼǻ󼼸�</div></th>
								<th><div>�߰��ǸŰ�</div></th>
								<th><div>�߰����ް�</div></th>
								<th><div>����</div></th>
							</tr>
							</thead>
							<tbody>
								<% for i=0 to oOptionMultipleType.FResultCount-1 %>
								<tr>
									<input type="hidden" name="TypeSeqTmp" value="<%= oOptionMultipleType.FItemList(i).FTypeSeq %>">
			        		<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>" width="30"><%= i+1 %></td>
			        		<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>">
			        	    <input type="text" class="formTxt" name="optionTypename" value="<%= oOptionMultipleType.FItemList(i).FoptionTypename %>" size="20" maxlength="32">
			        		</td>
			            <% TrFlag = false %>
					        	<% for k=0 to oOptionMultiple.FResultCount -1 %>
					        	<% if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(k).FoptionTypename) and (oOptionMultipleType.FItemList(i).FTypeSeq=oOptionMultiple.FItemList(k).FTypeSeq) then %>
					        	<% if (TrFlag) then %>
			        	</tr>
			       		<tr>
			           		<% end if %>
			            <input type="hidden" name="TypeSeq" value="<%= oOptionMultiple.FItemList(k).FTypeSeq %>">
			            <input type="hidden" name="KindSeq" value="<%= oOptionMultiple.FItemList(k).FKindSeq %>">
				        	<td><input type="text" class="formTxt" name="optionName" value="<%= oOptionMultiple.FItemList(k).FoptionKindName %>" size="20" maxlength="32"></td>
				        	<!-- <td></td> -->
				        	<td><input type="text" class="formTxt" name="optaddprice" value="<%= oOptionMultiple.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
				        	<td><input type="text" class="formTxt" name="optaddbuyprice" value="<%= oOptionMultiple.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
				        	<td><input type="button" class="button" value="����" onClick="DelItemOptionMultiple('<%= itemid %>','<%= oOptionMultiple.FItemList(k).FTypeSeq %>','<%= oOptionMultiple.FItemList(k).FKindSeq %>');" ></td>
			        	</tr>
				            <% pp = pp + 1 %>
				            <% TrFlag = true %>
			        		<% end if %>
			        		<% next %>
			    		<% next %>
			    	</tbody>
			    	</table>
	    	<% else %> <!-- ���Ͽɼ�  -->
	    		<table class="tbType1 listTb tMar10">
							<thead>
							<tr>
								<th><div>�ɼǱ��и�</div></th>
								<th><div>�ɼǻ󼼸�</div></th>
								<th><div>��뿩��</div></th>
								<th><div>ǰ������</div></th>
								<th><div>�߰��ǸŰ�</div></th>
								<th><div>�߰����ް�</div></th>
								<th><div>����</div></th>
							</tr>
							</thead>
							<tbody>
					    	 <tr>
				        	<td rowspan="<%= oitemoption.FResultCount %>">
				        	    <input type="text" class="formTxt" name="optionTypename" value="<%= oitemoption.FItemList(0).FoptionTypename %>" size="20" maxlength="32">
				        	</td>
				        	<% TrFlag = false %>
				        	<% for k=0 to oitemoption.FResultCount -1 %>
				        	<% if (TrFlag) then %>
				        </tr>
				        <tr bgcolor="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#FFFFFF","#DDDDDD") %>">
				            <% end if %>
				            <input type="hidden" name="itemoption" value="<%= oitemoption.FItemList(k).FItemOption %>">
				        	<td><input type="text" class="formTxt" name="optionName" value="<%= oitemoption.FItemList(k).FoptionName %>" size="20" maxlength="32"></td>
				        	<td><font color="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#000000","#FF0000") %>"><%= oitemoption.FItemList(k).Foptisusing %></font></td>
				        	<td><% if oitemoption.FItemList(k).IsOptionSoldOut then %><font color="red">ǰ��</font><% end if %></td>
				        	<td><input type="text" class="formTxt" name="optaddprice" value="<%= oitemoption.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
				        	<td><input type="text" class="formTxt" name="optaddbuyprice" value="<%= oitemoption.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
				        	<td><input type="button" class="button" value="����" onClick="DelItemOption('<%= itemid %>','<%= oitemoption.FItemList(k).Fitemoption %>');" ></td>
				        </tr>
				            <% pp = pp + 1 %>
				            <% TrFlag = true %>
				        	<% next %>
				        </tr>
				      </tbody>
				    </table>
    		<% end if %>
	<% end if %>
</form>  
<div class="tPad15 ct">
	<input type="button" value=" �ɼ� ���� ����" onclick="EditOptionInfo();" class="btn3 btnRd" /> 
</div>
			 

<form name="frmOption" method="post" action="do_waititemoptionedit.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
<input type="hidden" name="typeseq">
<input type="hidden" name="kindseq">
</form>
<%
set oitem = Nothing
set oOptionMultipleType = Nothing
set oOptionMultiple = Nothing
set oitemoption = Nothing
%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->