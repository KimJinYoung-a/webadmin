<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' Hieditor : ������ ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim mode
dim itemid, itemoption
dim oitem, oitemoption, oOptionMultipleType, oOptionMultiple

itemid = requestCheckvar(getNumeric(request("itemid")),10)
if itemid="" then itemid=0
mode= requestCheckvar(request("mode"),32)
itemoption= requestCheckvar(request("itemoption"),10)

dim sqlStr
dim ErrStr

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

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
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
if oitem.FOneItem.Fsellcash>0 then
	ItemDefaultMargin = 100-CLng(oitem.FOneItem.FBuycash/oitem.FOneItem.Fsellcash*100*100)/100
else
	ItemDefaultMargin = 0
end if

''20091126���� : �ٹ����ٹ��, �ɼ��� ���»�ǰ�� ����� ������� �ɼ��߰� �Ұ�
'20150821 �߰�����: �ٹ����ٹ�� ��� �ְų� �Ǹų��� �ִ� ��� �ɼǸ� ���� �����ڸ� �����ϵ���
'// 2021-04-08 : �԰� �Ǹ� ��� ���� ��� �ɼ��߰� ����
dim OptionAddDisable : OptionAddDisable = false
dim OptionModDisable : OptionModDisable = false

if (oitem.FOneItem.FMwDiv<>"U") then
    sqlStr =  " select isNull(sum(CNT),0) as CNT, isNull(sum(TotCnt),0) as TotCnt "
    sqlStr = sqlStr & " from ( "
    sqlStr = sqlStr & "     select case when itemoption = '0000' then count(itemid) else 0 end as CNT"
    sqlStr = sqlStr & " , count(itemid) as TotCnt "
    sqlStr = sqlStr & " from db_summary.dbo.tbl_current_logisstock_summary with (nolock)"
    sqlStr = sqlStr & " where itemgubun='10'"
    sqlStr = sqlStr & " and itemid="&itemid
    sqlStr = sqlStr & " and itemoption >= '0000' "
    sqlStr = sqlStr & " and (ipgono <> 0 or sellno <> 0 or toterrno <> 0 or totsysstock <> 0 or availsysstock <> 0 or realstock <> 0 or preorderno <> 0 or preordernofix <> 0) "
    sqlStr = sqlStr & "  group by itemid, itemoption "
    sqlStr = sqlStr & " ) as T "
    rsget.Open sqlStr,dbget,1
        OptionAddDisable = rsget("CNT")>0
        OptionModDisable = rsget("TotCnt")>0
        if  C_ADMIN_AUTH then
         ''OptionModDisable = False '��� ������ ������ �����ڵ� �����ϴ°��� ������
        end if

    rsget.Close

end if

'/�ӽ� �б� ó��
if itemid="1521739" then OptionAddDisable = false
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
var VItemDefaultMargin = <%= ItemDefaultMargin %>;
function EditOptionInfo(){
    var frm = document.frmEdit;
    var optAddpriceExists = false;
    console.log('����');
    if (frm.mode.value=="editOptionMultiple"){
        console.log('���߿ɼ�');
        //���߿ɼ�
        if (!frm.optionTypename.length){
            if (frm.optionTypename.value.length<1){
                alert('�ɼ� ���и��� �Է��ϼ���.');
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
        }else{
            for (var i=0;i<frm.optionName.length;i++){
                if (frm.optionName[i].value.length<1){
                    alert('�ɼǸ��� �Է��ϼ���.');
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
                alert('�߰��ݾ��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                frm.optaddprice.focus();
                return;
            }

            if(frm.optaddprice.value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
             }

            if (!IsDigit(frm.optaddprice.value)){
                <% if (C_ADMIN_AUTH) then %>
                if (!confirm('�߰��ݾ��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                    frm.optaddprice.focus();
                   return;
                }
                <% else %>
                alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                frm.optaddprice.focus();
                return;
                <% end if %>
            }

            if (Math.abs(frm.optaddbuyprice.value*1)>Math.abs(frm.optaddprice.value*1)) {
                alert('���ް��� �߰��ݾ� ���� Ŭ �� �����ϴ�.');
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
                    alert('�߰��ݾ��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }

                if(frm.optaddprice[i].value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }

                if (!IsDigit(frm.optaddprice[i].value)){
                    <% if (C_ADMIN_AUTH) then %>
                    if (!confirm('�߰��ݾ��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                        frm.optaddprice[i].focus();
                        return;
                    }
                    <% else %>
                    alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                    frm.optaddprice[i].focus();
                    return;
                    <% end if %>
                }

                if (Math.abs(frm.optaddbuyprice[i].value*1)>Math.abs(frm.optaddprice[i].value*1)) {
                    alert('���ް��� �߰��ݾ� ���� Ŭ �� �����ϴ�.');
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

        //�߰��ݾ�-���ް�
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('�߰��ݾ� ���ް��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                frm.optaddbuyprice.focus();
                return;
            }

            if(frm.optaddbuyprice.value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }


            if (!IsDigit(frm.optaddbuyprice.value)){
                <% if (C_ADMIN_AUTH) then %>
                if (!confirm('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                    frm.optaddbuyprice.focus();
                    return;
                }
                <% else %>
                alert('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.');
                frm.optaddbuyprice.focus();
                return;
                <% end if %>
              }

        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('�߰��ݾ� ���ް��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                    frm.optaddbuyprice[i].focus();
                    return;
                }

                if (!IsDigit(frm.optaddbuyprice[i].value)){
                   <% if (C_ADMIN_AUTH) then %>
                   if (!confirm('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                    <% else %>
                    alert('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                   <% end if%>
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
        console.log('���Ͽɼ�');
        //���Ͽɼ�
        if (frm.optionTypename.value.length<1){
            alert('�ɼ� ���и��� �Է��ϼ���.');
            frm.optionTypename.focus();
            return;
        }

        if (!frm.optionName.length){
            if (frm.optionName.value.length<1){
                alert('�ɼǸ��� �Է��ϼ���.');
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
                alert('�߰��ݾ��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                frm.optaddprice.focus();
                return;
            }



            //if (!IsDigit(frm.optaddprice.value)){
               <% 'if (C_ADMIN_AUTH) then %>
              //  if (!confirm('�߰��ݾ��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
              //      frm.optaddprice.focus();
              //      return;
              //  }
                <% 'else %>
                //alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                //frm.optaddprice.focus();
                //return;
                <% 'end if %>
            //}

            if (Math.abs(frm.optaddbuyprice.value*1)>Math.abs(frm.optaddprice.value*1)) {
                alert('���ް��� �߰��ݾ� ���� Ŭ �� �����ϴ�.');
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
                    alert('�߰��ݾ��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }

                if(frm.optaddprice[i].value<0){
                    alert('�߰��ݾ׿� ���̳ʽ� �ݾ��� �Է��Ҽ� �����ϴ�. (�߰��ݾ��� ������ 0)');
                    frm.optaddprice[i].focus();
                    return;
                }

                if (!IsDigit(frm.optaddprice[i].value)){
                    <% if (C_ADMIN_AUTH) then %>
                    if (!confirm('�߰��ݾ��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                        frm.optaddprice[i].focus();
                        return;
                    }
                    <% else %>
                    alert('�߰��ݾ��� ���ڸ� �����մϴ�.');
                    frm.optaddprice[i].focus();
                    return;
                    <% end if %>
                }

                if (Math.abs(frm.optaddbuyprice[i].value*1)>Math.abs(frm.optaddprice[i].value*1)) {
                    alert('���ް��� �߰��ݾ� ���� Ŭ �� �����ϴ�.');
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

        //�߰��ݾ�-���ް�
        if (!frm.optaddbuyprice.length){
            if (frm.optaddbuyprice.value.length<1){
                alert('�߰��ݾ� ���ް��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                frm.optaddbuyprice.focus();
                return;
            }

            if (!IsDigit(frm.optaddbuyprice.value)){
                <% if (C_ADMIN_AUTH) then %>
                if (!confirm('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                    frm.optaddbuyprice.focus();
                    return;
                }
                <% else %>
                alert('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.');
                frm.optaddbuyprice.focus();
                return;
                <% end if %>
            }
        }else{
            for (var i=0;i<frm.optaddbuyprice.length;i++){
                if (frm.optaddbuyprice[i].value.length<1){
                    alert('�߰��ݾ� ���ް��� �Է��ϼ���. (�߰��ݾ��� ������ 0)');
                    frm.optaddbuyprice[i].focus();
                    return;
                }

                if (!IsDigit(frm.optaddbuyprice[i].value)){
                    <% if (C_ADMIN_AUTH) then %>
                    if (!confirm('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.\r\n����Ͻðڽ��ϱ�?[�����ڸ��]')){
                        frm.optaddbuyprice[i].focus();
                        return;
                    }
                    <% else %>
                    alert('�߰��ݾ� ���ް��� ���ڸ� �����մϴ�.');
                    frm.optaddbuyprice[i].focus();
                    return;
                    <% end if %>
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
    console.log('�ɼǳ�');
    <% if (oitem.FOneItem.FMwDiv<>"U") then %>
    //�ٹ���� �ɼ� �߰��ݾ� ���Ұ� �ϰ� 20120326
    <% if NOT ((session("ssBctID")="icommang") or (session("ssBctID")="hrkang97")) then %>  //201509/01
    if (optAddpriceExists){
        alert('�ٹ����� ����� ��� �ɼ� �߰��ݾ��� ����� �� �����ϴ�.');
        return;
    }
    <% else %>
    if (optAddpriceExists){
        alert('������ ���� MODE.');

    }
    <% end if %>
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

	//������ �������� Ȯ��
	if($("input[name='TypeSeq']").last().val()>typeseq) {
		if($("input[name='TypeSeq'][value='"+typeseq+"']").length<=1) {
			alert("�� ������ �����ϴ� �� ������ �ɼǱ����� ������ �� �����ϴ�.");
			return;
		} else {
			alert("�� ������ �ɼǱ����� ������ �� ������ �����ϼ���.");
		}
	}

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
    <% if (OptionAddDisable) then %>
    	alert('�ٹ����� ��� �ɼ� ���»�ǰ�� �԰�(���ֹ� ����), �Ǹų��� �Ǵ� �������� �����Ƿ� �ɼ��߰� �Ұ��մϴ�.');
    	return;
    <% else %>
	    var popwin = window.open('pop_optionAdd.asp?itemid=' + iitemid,'pop_optionAdd','width=800,height=600,scrollbars=yes,resizable=yes');
	    popwin.focus();
    <% end if %>
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td valign="top" bgcolor="F4F4F4">
	        	<b>�ɼǼ���</b><br>

	        	<br>- �ɼ��� �߰� �Ǵ� �����Ҽ� �ֽ��ϴ�.
	        	<br>- �Ǹ�/�԰�/���� ������ �ִ� �ɼ��� ������ �Ұ����մϴ�.(������ ���� �����ϼ���)
	        	<%if (oitem.FOneItem.FMwDiv<>"U") then %>
	        	<br>- �Ǹ�/�԰�/���� ������ �ִ� �ɼ��� �ɼǸ� ������ ������ �Ǵ� ���ȹ���� �����ּ���
	        	<%end if%>
	        	<br>- �߰��ݾ��� ������� �ڵ����� ǥ�õ˴ϴ�.(�ɼ� �� �߰��ݾ��� ���� ������)
	        </td>
	</tr>
	</form>
</table>
<p>
<!-- ǥ ��ܹ� ��-->
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
  <form name=frmmaster method=post action=do_adminitemoptionedit.asp onSubmit="return false;">
  <input type="hidden" name="itemid" value="<%= oitem.FOneItem.Fitemid %>">
  <input type="hidden" name="mode" value="">
  <input type="hidden" name="arritemoption" value="">
  <input type="hidden" name="arritemoptionname" value="">
	<tr>
		<td width=120 height="25" bgcolor="#DDDDFF" align="center">��ǰ�ڵ�</td>
		<td  bgcolor="#FFFFFF"><%= itemid %></td>
		<td width=240 bgcolor="#DDDDFF" align="center">�ɼ� ���� �̸�����</td>
	</tr>
	<tr>
		<td width=120 height="25" bgcolor="#DDDDFF" align="center">��ǰ��</td>
		<td bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
		<td width=200 bgcolor="#FFFFFF" rowspan="2" align="center">
		<%= getOptionBoxHTML_FrontType(itemid) %>
		</td>
	</tr>
	<tr>
		<td width=120 height="25" bgcolor="#DDDDFF" align="center">�귣��</td>
		<td bgcolor="#FFFFFF"><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
	</tr>
  </form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmEdit" method="post" action="do_adminitemoptionedit.asp">
<input type="hidden" name="itemid" value="<%= itemid %>">
<% if (oitemoption.IsMultipleOption) then %>
<input type="hidden" name="mode" value="editOptionMultiple">
<% else %>
<input type="hidden" name="mode" value="editOption">
<% end if %>
	<tr height="25" bgcolor="FFFFFF">

		<td colspan="8">
		    <table width="100%" cellpadding="0" cellspacing="0" border="0" class="a" >
		    <tr>
		        <td>��ϵ� �ɼ� ����Ʈ</td>
		        <td width="80" align="right"><input type="button" class="button" value="�ɼ��߰� +" onClick="AddOptionPop('<%= itemid %>');"></td>
		    </tr>
		    </table>
		</td>
	</tr>
	<% if oitemoption.FResultCount<1 then %>
    <tr height="25" bgcolor="#FFFFFF">
	    <td colspan="8" align=center>��ϵ� �ɼ��� �����ϴ�.</td>
    </tr>
    <% else %>
        <% if (oitemoption.IsMultipleOption) then %>
        <!-- ���߿ɼ� -->
        <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="30">����</td>
        	<td width="200">�ɼǱ��и�</td>
        	<td >�ɼǻ󼼸�</td>
        	<!--
        	<td width="40">���<br>����</td>
        	-->
        	<td width="80">�߰�����</td>
        	<td width="80">���ް�</td>
        	<td width="80">����</td>
        </tr>
        <% for i=0 to oOptionMultipleType.FResultCount-1 %>
    	<tr align="center" bgcolor="#FFFFFF">
    	    <input type="hidden" name="TypeSeqTmp" value="<%= oOptionMultipleType.FItemList(i).FTypeSeq %>">
        	<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>" width="30"><%= i+1 %></td>
        	<td rowspan="<%= oOptionMultipleType.FItemList(i).FOptionCount %>">
        	    <input type="text" name="optionTypename" value="<%= oOptionMultipleType.FItemList(i).FoptionTypename %>" size="20" maxlength="20" <%if OptionModDisable  then%>readonly class="text_ro" <%else%>class="text" <%end if%>>
        	</td>
            <% TrFlag = false %>
        	<% for k=0 to oOptionMultiple.FResultCount -1 %>
        	<% if (oOptionMultipleType.FItemList(i).FoptionTypename=oOptionMultiple.FItemList(k).FoptionTypename) and (oOptionMultipleType.FItemList(i).FTypeSeq=oOptionMultiple.FItemList(k).FTypeSeq) then %>
        	<% if (TrFlag) then %>
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
            <% end if %>
            <input type="hidden" name="TypeSeq" value="<%= oOptionMultiple.FItemList(k).FTypeSeq %>">
            <input type="hidden" name="KindSeq" value="<%= oOptionMultiple.FItemList(k).FKindSeq %>">
        	<td><input type="text"  name="optionName" value="<%= oOptionMultiple.FItemList(k).FoptionKindName %>" size="20" maxlength="20" <%if OptionModDisable then%>readonly class="text_ro" <%else%>class="text" <%end if%>></td>
        	<!-- <td></td> -->
        	<td><input type="text" class="text" name="optaddprice" value="<%= oOptionMultiple.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
        	<td><input type="text" class="text" name="optaddbuyprice" value="<%= oOptionMultiple.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
        	<td><input type="button" class="button" value="����" onClick="DelItemOptionMultiple('<%= itemid %>','<%= oOptionMultiple.FItemList(k).FTypeSeq %>','<%= oOptionMultiple.FItemList(k).FKindSeq %>');" ></td>
        </tr>
            <% pp = pp + 1 %>
            <% TrFlag = true %>
        	<% end if %>
        	<% next %>
    	<% next %>
	    <% else %>
	    <!-- ���Ͽɼ�  -->
	    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="200">�ɼǱ��и�</td>
        	<td >�ɼǻ󼼸�</td>
        	<td width="40">���<br>����</td>
        	<td width="40">ǰ��<br>����</td>
        	<td width="80">�߰�����</td>
        	<td width="80">���ް�</td>
        	<td width="80">����</td>
        </tr>
	    <tr align="center" bgcolor="#FFFFFF">
        	<td rowspan="<%= oitemoption.FResultCount %>">
        	    <input type="text"  name="optionTypename" value="<%= oitemoption.FItemList(0).FoptionTypename %>" size="20" maxlength="20" <%if OptionModDisable  then%>readonly class="text_ro" <%else%>class="text" <%end if%>>
        	</td>
        	<% TrFlag = false %>
        	<% for k=0 to oitemoption.FResultCount -1 %>
        	<% if (TrFlag) then %>
        </tr>
        <tr align="center" bgcolor="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#FFFFFF","#DDDDDD") %>">
            <% end if %>
            <input type="hidden" name="itemoption" value="<%= oitemoption.FItemList(k).FItemOption %>">
        	<td><input type="text"  name="optionName" value="<%= oitemoption.FItemList(k).FoptionName %>" size="20" maxlength="40" <%if OptionModDisable then%>readonly class="text_ro" <%else%>class="text" <%end if%>></td>
        	<td><font color="<%= ChkIIF(oitemoption.FItemList(k).Foptisusing="Y","#000000","#FF0000") %>"><%= oitemoption.FItemList(k).Foptisusing %></font></td>
        	<td><% if oitemoption.FItemList(k).IsOptionSoldOut then %><font color="red">ǰ��</font><% end if %></td>
        	<td><input type="text" class="text" name="optaddprice" value="<%= oitemoption.FItemList(k).Foptaddprice %>" size="9" maxlength="9" style="text-align:right" onKeyUp="AutoCalcuBuyPrice(this,'<%= pp %>');"></td>
        	<td><input type="text" class="text" name="optaddbuyprice" value="<%= oitemoption.FItemList(k).Foptaddbuyprice %>" size="9" maxlength="9" style="text-align:right"></td>
        	<td><input type="button" class="button" value="����" onClick="DelItemOption('<%= itemid %>','<%= oitemoption.FItemList(k).Fitemoption %>');" ></td>
            <% pp = pp + 1 %>
        </tr>
            <% TrFlag = true %>
        	<% next %>
        </tr>
    	<% end if %>
	<% end if %>
</form>
</table>
<p>
<% if oitemoption.FResultCount>0 then %>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#FFFFFF>
<tr height="30">
    <td align="center"><input type="button" value="�ɼ� ���� ����" onClick="EditOptionInfo();"></td>
</tr>
</table>
<% end if %>

<form name="frmOption" method="post" action="do_adminitemoptionedit.asp">
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
