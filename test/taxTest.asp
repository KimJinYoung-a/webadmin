
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
dim dzConnURL : dzConnURL = "http://www.bill36524.com"
IF application("Svr_Info")="Dev" THEN
    dzConnURL = "http://www.bill36524.com:8090"
end if
    
function ShowDataSetEx(ds)
    dim ret
    Dim dt 
    For Each dt In ds
        Dim dr 
        For Each dr In dt
            Dim df 
            
            For Each df In dr
                ret = ret + "[" + df.Name + "]=[" + df.Value + "]" + vbCrLf
            Next
        
        Next
    Next
    
    ShowDataSetEx = ret
End function

function IsSuccessCMD(ds)
    IsSuccessCMD = (getDsValue(ds,"RESULT")="00000")
end function

function getDsValue(ds,varName)
    getDsValue = ""
    Dim dt 
    For Each dt In ds
        Dim dr 
        For Each dr In dt
            Dim df 
            
            For Each df In dr
                if (UCase(df.Name)=UCase(varName)) then
                    getDsValue = df.Value
                    Exit Function
                end if
            Next
        
        Next
    Next
    
end function


dim DzBill365Api, dzrow
dim ds, resultmsg
dim dr_tbtax

set DzBill365Api = Server.CreateObject("DzEBankSDK.DzBill365Api")
DzBill365Api.EnablePointPopup = 0
DzBill365Api.EnableLoginPopup = 0
DzBill365Api.ConnectionURL = dzConnURL	''8090 �׽�Ʈ ��Ʈ ,80 ������Ʈ :8090
		        
dim intRet : intRet = DzBill365Api.InitApi()

response.write "intRet="&intRet&"<br>"
   
IF (intRet<>"1") then
    response.write "Bill365Api�� �ʱ�ȭ ���� ���߽��ϴ�."
    set DzBill365Api = Nothing
    response.end
end if

set dzrow = Server.CreateObject("DzEBankSDK.DzDataRow")
dzrow.AddNew "ID", "tenbyten" 
dzrow.AddNew "PASSWD", "tenbyten" ''"20011010" 


''response.end
''response.write DzBill365Api.Islogon&"<br>"

set ds = DzBill365Api.Login(dzrow)  '''���� CUP up...;; ==> ���� �������α׷����� ���.
		
	
if (Not IsSuccessCMD(ds)) then
    resultmsg = getDsValue(ds,"RESULT_MSG")
else
    
end if

set dzrow = Nothing
set ds = Nothing

if (resultmsg<>"") then
    response.write "bill36524 �α��ο� ���� �Ͽ����ϴ�." + resultmsg
    set DzBill365Api = Nothing
    response.end
end if
response.end

set dr_tbtax = Server.CreateObject("DzEBankSDK.DzDataRow")

call dr_tbtax.AddNew("FG_BILL","1")     ''//û��1 ����2
call dr_tbtax.AddNew("YN_TURN","N")     ''//Y������ N������
call dr_tbtax.AddNew("FG_IO","2")       ''//1���� 2����                            
call dr_tbtax.AddNew("FG_PC","1")       ''//1��� 2����                            
call dr_tbtax.AddNew("FG_FINAL","1")    ''//0���� 1 �߼� 2���� 3�ݷ� 4������ҿ�û 
call dr_tbtax.AddNew("YN_CSMT","N")     ''//N������� Y����Ź����                  
                                                                                   
call dr_tbtax.AddNew("FG_VAT","1")      ''//1���� 2���� 3�鼼                      
call dr_tbtax.AddNew("AM","10")         ''//���ް���                               
call dr_tbtax.AddNew("AM_VAT","1")      ''//�ΰ���                                 
call dr_tbtax.AddNew("AMT","11")        ''//�հ�                                   
call dr_tbtax.AddNew("AMT_CASH","0")    ''//����                                  
call dr_tbtax.AddNew("AMT_CHECK","0")   ''//��ǥ                                   
call dr_tbtax.AddNew("AMT_NOTE","0")    ''//����                                   
                                                        ''������(�Ǹ���)
call dr_tbtax.AddNew("YMD_WRITE","20090722")
call dr_tbtax.AddNew("SELL_NO_BIZ","1208615854") ''6666666666
call dr_tbtax.AddNew("SELL_NM_CORP","�������")
call dr_tbtax.AddNew("SELL_NM_CEO","������")
call dr_tbtax.AddNew("SELL_ADDR1","aa")
call dr_tbtax.AddNew("SELL_DAM_DEPT","dep")
call dr_tbtax.AddNew("SELL_DAM_NM","��")
call dr_tbtax.AddNew("SELL_DAM_EMAIL","")
                                                        ''������
call dr_tbtax.AddNew("BUY_NO_BIZ","9999999999")
call dr_tbtax.AddNew("BUY_NM_CEO","������")
call dr_tbtax.AddNew("BUY_NM_CORP","����ȸ��")
call dr_tbtax.AddNew("BUY_DAM_EMAIL","")

''//���ݰ�꼭 ǰ�� : 4�� ���� ����

dim dr_tbtaxline1 
set dr_tbtaxline1 = Server.CreateObject("DzEBankSDK.DzDataRow")

call dr_tbtaxline1.AddNew("ITEM_STD", "Item")
call dr_tbtaxline1.AddNew("NM_ITEM", "Item")
call dr_tbtaxline1.AddNew("NO_ITEM", "01")
call dr_tbtaxline1.AddNew("AM", "10")
call dr_tbtaxline1.AddNew("AM_VAT", "1")
call dr_tbtaxline1.AddNew("AMT", "11")
call dr_tbtaxline1.AddNew("DD_WRITE", "22")
call dr_tbtaxline1.AddNew("MM_WRITE", "07")

dim dr_tbtaxline2 
set dr_tbtaxline2 = Server.CreateObject("DzEBankSDK.DzDataRow")

call dr_tbtaxline2.AddNew("ITEM_STD", "Item")
call dr_tbtaxline2.AddNew("NM_ITEM", "Item")
call dr_tbtaxline2.AddNew("NO_ITEM", "01")
call dr_tbtaxline2.AddNew("AM", "10")
call dr_tbtaxline2.AddNew("AM_VAT", "1")
call dr_tbtaxline2.AddNew("AMT", "11")
call dr_tbtaxline2.AddNew("DD_WRITE", "22")
call dr_tbtaxline2.AddNew("MM_WRITE", "07")            

dim dt_tbtaxline
set dt_tbtaxline = Server.CreateObject("DzEBankSDK.DzDataTable")
call dt_tbtaxline.Add(dr_tbtaxline1)
call dt_tbtaxline.Add(dr_tbtaxline2)
''//call dt_tbtaxline.Add("dr_tbtaxline3", dr_tbtaxline3)

set ds = DzBill365Api.SendTaxAccount(dr_tbtax, dt_tbtaxline)

response.write ShowDataSetEx(ds)

set dr_tbtaxline1 = Nothing
set dr_tbtaxline2 = Nothing
set dt_tbtaxline = Nothing
set dr_tbtax = Nothing
set ds = Nothing
set DzBill365Api = Nothing
%>