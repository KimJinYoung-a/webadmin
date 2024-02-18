<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : index.asp
' Discription : �³�Ʈ ���̾ ��ƼĿ ������ ��� â
' History : 2023.03.31 ������
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/goodNoteDiaryCls.asp"-->
<%
dim menupos, iCurrpage, iPageSize, iPerCnt, iTotCnt
dim arrList, cEvtList, ix
	menupos = request("menupos")
	iCurrpage = Request("iC")	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����

	'������ ��������
	set cEvtList = new GoodNoteDiaryCls
		cEvtList.FCurrPage = iCurrpage	'����������
		cEvtList.FPageSize = iPageSize '���������� ���̴� ���ڵ尹��
 		arrList = cEvtList.getStickerList	'�����͸�� ��������
 		iTotCnt = cEvtList.FTotalCount	'��ü ������  ��
 	set cEvtList = nothing
%>
<script>
function TnTrainThemeItemBannerReg(){
    var winpop = window.open("/admin/eventmanage/goodnote/pop_sticker_register.asp","winpop","width=1200,height=800,scrollbars=yes,resizable=yes");
    winpop.focus();
}
function fnStickerEdit(idx){
    var winEditpop = window.open("/admin/eventmanage/goodnote/pop_sticker_register.asp?idx="+idx,"winpop","width=1200,height=800,scrollbars=yes,resizable=yes");
    winEditpop.focus();
}
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>�³�Ʈ ���̾ ��ƼĿ ����</h1>
	</div>
    <button class="btn4 btnBlock btnWhite2 tMar10 tPad20 bPad20 lt" onClick="TnTrainThemeItemBannerReg();return false;"><span class="mdi mdi-plus cBl4 fs15"></span> ��ƼĿ �߰�</button>
    <% If isArray(arrList) Then %>
    <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
        <table class="tableV19A tableV19B tMar10">
            <thead>
                <tr>
                    <th>No</th>
                    <th>����</th>
                    <th>������</th>
                    <th>��뿩��</th>
                    <th>�����</th>
                </tr>
            <thead>
            <tbody>
                <% For ix = 0 To UBound(arrList,2) %>
                <tr onclick="fnStickerEdit(<%=arrList(0,ix)%>);">
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="mdi fs20"><%=arrList(0,ix)%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="mdi fs20"><%=arrList(1,ix)%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="previewThumb50W"><%=FormatDate(arrList(2,ix),"0000.00.00")%>~<%=FormatDate(arrList(3,ix),"0000.00.00 00:00:00")%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><%=arrList(4,ix)%></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><%=FormatDate(arrList(5,ix),"0000.00.00 00:00:00")%></td>
                </tr>
                <% Next %>
            </tbody>
        </table>
    </div>
    <% End If %>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->