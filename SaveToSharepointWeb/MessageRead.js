(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      //var element = document.querySelector('.ms-MessageBanner');
      //messageBanner = new fabric.MessageBanner(element);
      //messageBanner.hideBanner();
        //loadProps();
        getAccessToken();
    });
  };

    function getSiteCollections(ssotoken) {

        $.ajax({
            type: "GET",
            url: "api/GetSiteCollections",
            headers: {
                "Authorization": "Bearer " + ssotoken
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the sitecollection data");
            console.log(data);
        }).fail(function (error) {
            console.log("Fail to fetch site collections");
            console.log(error);
        });
    }

    function getDocumentLibraries(ssotoken) {
        var siteinfo = {
            Id: "vssworks.sharepoint.com,40aea85f-caf1-47f5-838c-19a2be1a64fa,b4a807ac-64fe-4b47-b02c-a18686f9551d",
            Name:"XRMTest"
        }

        $.ajax({
            type: "POST",
            url: "api/GetDocumentLibraries",
            headers: {
                "Authorization": "Bearer " + ssotoken
            },
            contentType: "application/json; charset=utf-8",
            data:JSON.stringify(siteinfo)
        }).done(function (data) {
            console.log("Fetched the sitecollection data");
            console.log(data);
        }).fail(function (error) {
            console.log("Fail to fetch site collections");
            console.log(error);
        });
    }

    function getDocumentLibrariesFolder(ssotoken) {
        var documentinfo = {
            Id: "b!X6iuQPHK9UeDjBmivhpk-qwHqLT-ZEdLsCyhhob5VR2Gw1vmxNYGRYQQLXGdBwfG",
            Name: "Contracts"
        }

        $.ajax({
            type: "POST",
            url: "api/GetLibraryFolders",
            headers: {
                "Authorization": "Bearer " + ssotoken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(documentinfo)
        }).done(function (data) {
            console.log("Fetched the sitecollection data");
            console.log(data);
        }).fail(function (error) {
            console.log("Fail to fetch site collections");
            console.log(error);
        });
    }

    function saveAttachment(ssotoken) {
        var attachmentinfo = {
            messageId: "AAMkAGFmMTIzMGI2LWYzYjItNGRmNi1iNDA0LWE4ZjQ2ZmE3MWFmYQBGAAAAAACSvln3nLsZRbHnz5sHZ99OBwBs3XOIGghdTKdvFU-RrpirAAAAAAEMAABs3XOIGghdTKdvFU-RrpirAAAOr6LAAAA=",
            driveId: "b!X6iuQPHK9UeDjBmivhpk-qwHqLT-ZEdLsCyhhob5VR2Gw1vmxNYGRYQQLXGdBwfG",
            attachmentIds: ["AAMkAGFmMTIzMGI2LWYzYjItNGRmNi1iNDA0LWE4ZjQ2ZmE3MWFmYQBGAAAAAACSvln3nLsZRbHnz5sHZ99OBwBs3XOIGghdTKdvFU-RrpirAAAAAAEMAABs3XOIGghdTKdvFU-RrpirAAAOr6LAAAABEgAQALAGpl5MtwBItN4zA8SRnfc="]
        }

        $.ajax({
            type: "POST",
            url: "api/SaveAttachments",
            headers: {
                "Authorization": "Bearer " + ssotoken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(attachmentinfo)
        }).done(function (data) {
            console.log("Fetched the sitecollection data");
            console.log(data);
        }).fail(function (error) {
            console.log("Fail to fetch site collections");
            console.log(error);
        });
    }

    function getAccessToken() {
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync(function (result) {
                if (result.status === "succeeded") {
                    console.log("token was fetched ");
                    //getSiteCollections(result.value);
                    saveAttachment(result.value);
                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ forceConsent: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                           // getSiteCollections(result.value);
                            saveAttachment(result.value);
                        }
                        else {
                            console.log("No token was fetched " + result.error.code);
                            //getSiteCollections();
                        }
                    });
                }
                else {
                    console.log("error while fetching access token " + result.error.code);
                }
            });
        }
    }

    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  
    function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();