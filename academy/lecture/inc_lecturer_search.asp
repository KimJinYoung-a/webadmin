<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ� ���� ���ڵ� �˻� ,������ ���Ͽ�, ��ũ�� ���� �����ϱ� ���� ����
' History : 2011.02.25 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<%
dim OrderSerial ,osearch , userid , sql , RowNum , itemid ,oroom , barcoderoomid
dim startdate,entryname ,lec_title ,SubTotalPrice ,barcodelecprice ,barcodematprice
	OrderSerial = RequestCheckvar(request("OrderSerial"),16)
	itemid = RequestCheckvar(request("itemid"),10)
	startdate = RequestCheckvar(request("startdate"),10)
	entryname = RequestCheckvar(request("entryname"),32)
	lec_title = request("lec_title")
	SubTotalPrice = RequestCheckvar(request("SubTotalPrice"),10)
	barcodelecprice = RequestCheckvar(request("barcodelecprice"),10)
	barcodematprice = RequestCheckvar(request("barcodematprice"),10)
	
	RowNum = 0
  	if lec_title <> "" then
		if checkNotValidHTML(lec_title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
if OrderSerial = "" or itemid = "" then
	response.write "<script>alert('�ֹ� ��ȣ�� ���ų� ���¹�ȣ�� �����ϴ�.\n������ ���� �ϼ���');</script>"
end if

set osearch = new CLectureFingerOrder
	osearch.frectOrderSerial = OrderSerial
	osearch.flecturer_search()
	
	'//�ֹ���ȣ�� �ش� ȸ���� ���̵� �����´�
	if osearch.ftotalcount > 0 then
		userid = osearch.FOneItem.fuserid
	else
		response.write "<script>alert('�ֹ������� �����ϴ�.������ ���� �ϼ���');</script>"
		response.end
	end if

	sql = ";WITH TMP_LIST AS"
	sql = sql & " (" 
	sql = sql & " select ROW_NUMBER() OVER (ORDER BY o.lecstartdate) AS RowNum,"
	sql = sql & " d.orderserial, o.*"  
	sql = sql & " from db_academy.dbo.tbl_academy_order_master m"
	sql = sql & " Join [db_academy].[dbo].tbl_academy_order_detail d"
	sql = sql & " on m.orderserial=d.orderserial" 
	sql = sql & " and m.ipkumdiv>3"
	sql = sql & " and m.cancelyn='N'"
	sql = sql & " left Join [db_academy].dbo.tbl_lec_item_option o"
	sql = sql & " on d.itemid=o.lecIdx"
	sql = sql & " and d.itemoption=o.lecOption"
	sql = sql & " where m.userid='"&userid&"'"
	sql = sql & " )"
	
	sql = sql & " select RowNum from TMP_LIST"
	sql = sql & " where orderserial='"&OrderSerial&"'"
	
	'response.write sql &"<br>"
	rsACADEMYget.Open sql, dbACADEMYget, 1
	
	'//�ش� ȸ���� �ΰ��� ���� Ƚ���� �����´�
	if Not rsACADEMYget.Eof then		
		RowNum = rsACADEMYget("RowNum")
	else
		RowNum = 1	
	end if
	
	rsACADEMYget.Close
	
	'//��ȸ���� ��� ����Ƚ�� 1 ����
	if userid = "" then RowNum = 1

set oroom = new CLectureFingerOrder
	oroom.frectitemid = itemid
	oroom.flecturer_room()
	
	'/�ش� ������ ���ǽ��� �����´�
	barcoderoomid = oroom.FOneItem.barcoderoomid
%>

	<script language="javascript">

	function DrawReceiptPrintobj_TEC(elementid,printname){
	    var objstring = "";
	    var e;
	    objstring = '<OBJECT name="' + elementid + '" ';
	    objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
	    objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
	    objstring = objstring + ' width=0 ';
	    objstring = objstring + ' height=0 ';
	    objstring = objstring + ' align=center ';
	    objstring = objstring + ' hspace=0 ';
	    objstring = objstring + ' vspace=0 ';
	    objstring = objstring + ' > ';
	    objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
	    objstring = objstring + ' </OBJECT>';
	    
	    document.write(objstring);
	}

	DrawReceiptPrintobj_TEC("TEC_DO2","TEC B-SV4");

	var OrderSerial = '<%= OrderSerial %>';
	var startdate = '<%= startdate %>';
	var entryname = '<%= entryname %>';
	var lec_title = '<%= lec_title %>';
	var SubTotalPrice = '<%= SubTotalPrice %>';
	var barcodelecprice = '<%= barcodelecprice %>';
	var barcodematprice = '<%= barcodematprice %>';
	var RowNum = '<%= RowNum %>';
	var barcoderoomid = '<%= barcoderoomid %>';

	var X = 7.9; //1.5;
	var Y = 8; //1.5;
	var F = 1; //1.4;
	
	if (TEC_DO2.IsDriver == 1){
		//TEC_DO2.SetCutter(1, 0, 0, 0 );
		//TEC_DO2.SetDriverStock(0, 1, 5, 0, 1);  // Default �����Ѱ���.
	   TEC_DO2.SetPaper(575,1450);
	   
       TEC_DO2.OffsetX = 0;
       TEC_DO2.OffsetY = 0;
       
       TEC_DO2.PrinterOpen();
       //TEC_DO2.SetDriverStock(0, 1, 2.0, 0, 3);
       //left
              
       TEC_DO2.PrintText(13*X, 48*Y, "arial", 22*F, 0, 0, OrderSerial);		//�ֹ���ȣ
              
       TEC_DO2.PrintText(100*X, 15*Y, "arial", 22*F, 0, 0, entryname);		//�̸�
       
       TEC_DO2.PrintText(100*X, 21*Y, "arial", 22*F, 0, 0, startdate);		//��������
       
       TEC_DO2.PrintText(93*X, 29*Y, "HY�߰��", 16*F, 0, 0, lec_title);		//���¸�
       
       TEC_DO2.PrintText(108*X, 34*Y, "arial", 22*F, 0, 0, barcoderoomid);		//���ǽ�
       
       TEC_DO2.PrintText(93*X, 44*Y, "arial", 22*F, 0, 0, SubTotalPrice);		//�հ�
       
       TEC_DO2.PrintText(93*X, 51*Y, "arial", 22*F, 0, 0, barcodelecprice);		//������
       
       TEC_DO2.PrintText(120*X, 51*Y, "arial", 22*F, 0, 0, barcodematprice);		//����
              
       TEC_DO2.PrintText(56*X, 30*Y, "arial", 22*F, 0, 0, RowNum);		//����Ƚ��
       
       //right �ֹ�����
             
       TEC_DO2.PrinterClose();
	}else {
		alert('TEC B-EV4-G ����̹��� ��ġ�� �ּ���')		
	}
	</script>
<%
	set osearch = nothing
	set oroom = nothing
%>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->