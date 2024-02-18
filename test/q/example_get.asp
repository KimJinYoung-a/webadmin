<%@ Language="VBScript" %>

<!DOCTYPE html>
<html lang="ko">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <title>Classic ASP RESTful API Client Get Example</title>
    </head>
    <body>
        <!--#include file="json.asp"-->
        <!--#include file="hmac.asp"-->
        <!--#include file="buildQuery.asp"-->
        <%

            'Remote JSON Request
            dim url, path, method, params
            dim access_key, secret_key, vendorId
            dim authorization

            access_key = "0af06fb7-3deb-4ac3-9a84-6d409a26d831"
            secret_key = "5474f1108ac5631e5977d4a6b7a6387426533582"
            vendorId = "A00039305"

'            set hash = CreateObject ("Scripting.Dictionary")
'            hash.add "categoryName", "???"

'            path = "/v2/providers/openapi/apis/api/v4/vendors/"&vendorId&"/returnRequests/234711033"
'            path = "/v2/providers/openapi/apis/api/v4/vendors/"&vendorId&"/returnRequests"
            path = "/v2/providers/seller_api/apis/api/v1/marketplace/meta/category-related-metas/display-category-codes/77723"
            url = "https://api-gateway.coupang.com" & path
            method = "GET"

            response.write url & "<br/>"

            response.write method & path & "<br/>"
'params = "searchType=timeFrame&createdAtFrom=2020-06-13T00:00&createdAtTo=2020-06-13T23:00&cancelType=CANCEL"
'params = "searchType=timeFrame&createdAtFrom=2020-06-13T00:00&createdAtTo=2020-06-14T00:00&cancelType=CANCEL"
            authorization = generateHmac(path, method, params, access_key, secret_key)

            response.write "qweqwe : " & authorization
            response.write "<br/>"
'response.end
            set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
            req.open method, url & "?" & params, false
            req.setRequestHeader "Authorization", authorization
            req.setRequestHeader "X-Requested-By", vendorId
            req.send ""

' ---------------------------------------------------------------------------

            dim myJSON
            set myJSON = JSON.parse(req.responseText)

            response.write req.status
            response.write "<br/>"
            response.write myJSON.message
            response.write "<br/>"

            response.write req.responseText

        %>
    </body>
</html>