<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/cscenter/action/incPaycoCommon.asp"-->
<%

''ÇÑ±Û ÇÑ±Û

Dim orderserial : orderserial=requestCheckvar(request("orderserial"),20)
Dim tid         : tid=requestCheckvar(request("tid"),100)
Dim paddparam   : paddparam=requestCheckvar(request("paddparam"),200)
Dim i, rURi
dim sellerKey, sellerOrderReferenceKey, orderNo, tmpArr

orderNo = tid
sellerOrderReferenceKey = "10x10"

sellerKey = Payco_sellerKey
tmpArr = Split(paddparam, "|")
if (UBound(tmpArr) = 1) then
	Select Case tmpArr(1)
		Case "WEB"
			sellerKey = Payco_sellerKey_WEB
		Case "MOB"
			sellerKey = Payco_sellerKey_MOB
		Case "APP"
			sellerKey = Payco_sellerKey_APP
		Case Else
			''
	End Select
end if


rURi = Payco_URL_bill & "/seller/receipt/" & sellerKey & "/" & sellerOrderReferenceKey & "/" & orderNo
Response.Redirect rURi

%>
<style>
body, tr, td {font-size:9pt; font-family:±¼¸²,verdana; color:#433F37; line-height:19px;}
table, img {border:none}

/* Padding ******/
.pl_01 {padding:1 10 0 10; line-height:19px;}
.pl_03 {font-size:20pt; font-family:±¼¸²,verdana; color:#FFFFFF; line-height:29px;}

/* Link ******/
.a:link  {font-size:9pt; color:#333333; text-decoration:none}
.a:visited { font-size:9pt; color:#333333; text-decoration:none}
.a:hover  {font-size:9pt; color:#0174CD; text-decoration:underline}

.txt_03a:link  {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:visited {font-size: 8pt;line-height:18px;color:#333333; text-decoration:none}
.txt_03a:hover  {font-size: 8pt;line-height:18px;color:#EC5900; text-decoration:underline}

.buttoncss {
	font-family: "Verdana", "µ¸¿ò";
	font-size: 9pt;
	background-color: #E6E6E6;
	border: 1px outset #BABABA;
	color: #000000;
	height: 20px;
	cursor:hand;
}

</style>
<body bgcolor="#FFFFFF" text="#242424" leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 bottommargin=0 rightmargin=0><center>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp" -->
