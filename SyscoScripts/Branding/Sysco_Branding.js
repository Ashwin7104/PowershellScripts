var siteURL ="https://sysco.sharepoint.com/sites/Communications_UAT";
function loadDOM()
{

$.ajax({
  url : _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid(" + _spPageContextInfo.userId+ ")",
  contentType : "application/json;odata=verbose",
   headers: {"accept": "application/json; odata=verbose"},
   success: function (data) {
   var loginName = data.d.Title;
    var checkAdmin=CheckIfUserCanView('CommCSadministrators')
   // GetAllGroups(loginName,checkAdmin);
   if(!checkAdmin)
   {
     GetAllGroups(loginName);
    }
       else
        {
        $("#userName").prepend("<p style='font-size:0.9em'>Hello, "+loginName+"</p>");
        }
},
   error: function (data) {
    console.log(data);
            }
  });

 
var pathname = window.location.pathname; // Returns path only
if(pathname.toLowerCase().indexOf("/sitepages/") > -1 || pathname.toLowerCase().indexOf("/pages/") > -1)
{
$("#siteIcon").append("<div class='siteHeading'>HR Secure SharePoint Site</div>");
$("#siteIcon").addClass('syscoSiteIcon');
$(".ms-mpSearchBox").addClass('syscoBanner');
$("#s4-titlerow").addClass('syscoBannerDim');
$("#titleAreaRow").addClass('syscoBannerDim');
$("#pageTitle").addClass('syscoHide');
$(".ms-breadcrumb-top").addClass('syscoHide');
$(".ms-listMenu-editLink").addClass('syscoHide');
$("#s4-titlerow").append("<div id='Syscobreadcrumb'><div id='userName'></div><div id='search'></div></div>");
var $button = $('#SearchBox').clone();
$('#SearchBox').remove();
$('#search').append($button);

 }
}

$( document ).ready(function() {
$(".ms-core-listMenu-item:contains('Recent')").parent().hide();
 SP.SOD.executeFunc('sp.js', 'SP.ClientContext', loadDOM);
});
function GetAllGroups(loginName) {
        var grpName = "";
		var clientContext = new SP.ClientContext(siteURL);
		var groups = clientContext.get_web().get_siteGroups()
		clientContext.load(groups,"Include(CanCurrentUserViewMembership,Title)");
		clientContext.executeQueryAsync(
		function(sender,args){
    	var groupIterator = groups.getEnumerator();
    	var myGroups = [];
    	while(groupIterator.moveNext()){
        var current = groupIterator.get_current();
        var isMemberOfGroup = current.get_canCurrentUserViewMembership() ;
        if(isMemberOfGroup){
        var flag='';
        
        var groupName =current.get_title();
        console.log(groupName);
        if (groupName.toLowerCase().startsWith("commgrp") && groupName != 'CommCSadministrators') {
                    var start_pos='';
                    var end_pos ='';
                    var start='';
                    if(groupName.indexOf('_')>-1)
                    {
                	start_pos = groupName.indexOf('_') + 1;
                	if(groupName.substring(start_pos,groupName.length).indexOf('_')>-1)
                	{
					end_pos = groupName.indexOf('_',start_pos);
					}
					else
					{
					end_pos=groupName.length;
					}
					}
					flag =groupName.substring(start_pos,end_pos);
					if(flag.indexOf('-')>-1)
					{
					start = flag.indexOf('-') + 1;
					flag =flag.substring(start,flag.length);												
                   }
        $("#userName").prepend("<p style='font-size:0.9em'>Hello, "+loginName+" ("+flag+")</p>");
        break;
        }            
    	}
    	}
   		},function(sender,args){console.log(args.get_message());});	
        }
     
    function CheckIfUserCanView(groupName) {
    var flag='';
        $.ajax({
            url: siteURL  + "/_api/web/sitegroups/getbyname('" + groupName + "')/CanCurrentUserViewMembership",
            headers: { Accept: "application/json;odata=verbose" },
            async: false,
            success: function (data) {
            if (data.d.CanCurrentUserViewMembership === true) {
                                      flag='exists';	
                                      }			
 			},
            error: function (data) {
               console.log(data);
               flag='';
            }
        });
return flag ;
    }

