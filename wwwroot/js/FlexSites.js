var users = [];

$('.flexdatalist').flexdatalist({
    minLength: 0,
    valueProperty: 'description',
    selectionrequired: false,
    visibleProperties: ["description"],
    searchIn: 'description', 
    multiple: true,
    searchContain: true,
    selectionrequired: 1,
    /*data: '/ProjectManagement/GetAllSites'*/
    data: '/WebView.Test2/ClientSpec/RTSM-TK5/ProjectManagement/GetAllSites'
   
   
});
