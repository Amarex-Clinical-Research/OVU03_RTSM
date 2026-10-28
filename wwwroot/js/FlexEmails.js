////var users = [];
////$('.flexdatalist').flexdatalist({
////    minLength: 1

////});

var users = [];

$('.flexdatalist').flexdatalist({
    minLength: 0,
    valueProperty: 'Title',
    visibleProperties: ["description"],
    searchIn: 'description',

    multiple: true,
    searchContain: true,

    /* data: '/ProjectManagement/GetAllEmails'*/
    data: AppPath + 'ProjectManagement/GetAllEmails',
    //data: '/ProjectManagement/GetAllEmails',
    //data: '/WebView.Test2/ClientSpecRTSM/RTSM_PEN/ProjectManagement/GetAllEmails'

});