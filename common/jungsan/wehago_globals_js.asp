/**
 * Wehago Object
 */

 ;var wehago_globals = {
    service_code: "<%= wehagoServiceCode %>",   // 발급받은 코드
    <% IF application("Svr_Info")="Dev" THEN %>
        //mode : "dev",   // dev-개발, live-운영
        mode : "live",   // dev-개발, live-운영
    <% else %>
        mode : "live",   // dev-개발, live-운영
    <% end if %>
};