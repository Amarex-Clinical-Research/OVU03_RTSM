////var users = [];

////$('.flexdatalist').flexdatalist({
////    minLength: 0,
////   /* valueProperty: 'Title',*/
////    visibleProperties: ["Title"],
////    searchIn: 'Title',

////    multiple: true,
////    searchContain: true,

////    valueSeparator: ';',
////    data: '/ProjectManagement/GetAllEmails',
////    /*data: '/WebView.Test2/ClientSpec/IRTRoadMap1/ProjectManagement/GetAllEmails'*/
    

////});
var users = [];

$('.flexdatalist').flexdatalist({
    minLength: 0,
    valueProperty: 'Title', //Value and description can not be different for it to be able to display prefilled values not in the list
    //// selectionrequired: false, //I think this causes it to only allow a value to be from the list and when false it allows both selecting or just entering information not in the list but it cant fill the field with data passed in from the model
   // visibleProperties: ["title"],
   // searchIn: 'title', //Name
    visibleProperties: ["description"],
    searchIn: 'description',
    //  groupBy: 'FromTime', // sponsor
    multiple: true,
    searchContain: true,
    /*data: '/ProjectManagement/GetAllEmails',*/ //live version
    // data: '/Payments/GetAllEmails', //local version
    valuesSeparator: ','
});

