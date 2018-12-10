(function () {
    "use strict";

    var messageBanner;
    var ssoToken;



    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $(".loader").css("display", "block");
            getAccessToken();

            $("#sitecollections").change((event) => {
                $("#drives").css("display", "block");
                getDocumentLibraries();
            });

            $("#drivesselect").change((event) => {
                $("#folderselect").css("display", "block");
                getDocumentLibrariesFolder();
                loadProps();
            });

        });
    };

    function getSiteCollections(token) {

        $.ajax({
            type: "GET",
            url: "api/GetSiteCollections",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the sitecollection data");
            $.each(data, (index, value) => {
                $("#sitecollections").append('<option value="' + value.Id + '">' + value.Name + '</option>');
            });
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch site collections");
            console.log(error);
        });
    }

    function getDocumentLibraries() {
        $(".loader").css("display", "block");
        var siteinfo = {
            Id: $("#sitecollections").find("option:selected").val(),
            Name: $("#sitecollections").find("option:selected").text()
        }

        $.ajax({
            type: "POST",
            url: "api/GetDocumentLibraries",
            headers: {
                "Authorization": "Bearer " + ssoToken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(siteinfo)
        }).done(function (data) {
            console.log("Fetched the document libraries data");
            $.each(data, (index, value) => {
                $("#drivesselect").append('<option value="' + value.Id + '">' + value.Name + '</option>');
            });
            $(".loader").css("display", "none");

        }).fail(function (error) {
            console.log("Fail to fetch document libraries");
            console.log(error);
        });
    }

    function getDocumentLibrariesFolder() {
        $(".loader").css("display", "block");
        var documentinfo = {
            Id: $("#drivesselect").find("option:selected").val(),
            Name: $("#drivesselect").find("option:selected").text()
        }

        $.ajax({
            type: "POST",
            url: "api/GetLibraryFolders",
            headers: {
                "Authorization": "Bearer " + ssoToken
            },
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(documentinfo)
        }).done(function (data) {
            console.log("Fetched the folders data");
            $.each(data, (index, value) => {
                $("#libraryfolders").append('<option value="' + value.Id + '">' + value.Name + '</option>');
            });

            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch the library folders");
            console.log(error);
        });
    }

    function saveAttachment() {
        var attachmentinfo = {
            messageId: "AAMkAGFmMTIzMGI2LWYzYjItNGRmNi1iNDA0LWE4ZjQ2ZmE3MWFmYQBGAAAAAACSvln3nLsZRbHnz5sHZ99OBwBs3XOIGghdTKdvFU-RrpirAAAAAAEMAABs3XOIGghdTKdvFU-RrpirAAAOr6LAAAA=",
            driveId: "b!X6iuQPHK9UeDjBmivhpk-qwHqLT-ZEdLsCyhhob5VR2Gw1vmxNYGRYQQLXGdBwfG",
            attachmentIds: ["AAMkAGFmMTIzMGI2LWYzYjItNGRmNi1iNDA0LWE4ZjQ2ZmE3MWFmYQBGAAAAAACSvln3nLsZRbHnz5sHZ99OBwBs3XOIGghdTKdvFU-RrpirAAAAAAEMAABs3XOIGghdTKdvFU-RrpirAAAOr6LAAAABEgAQALAGpl5MtwBItN4zA8SRnfc="]
        }

        $.ajax({
            type: "POST",
            url: "api/SaveAttachments",
            headers: {
                "Authorization": "Bearer " + ssoToken
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
                    ssoToken = result.value;
                    getSiteCollections(result.value);

                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ forceConsent: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                            ssoToken = result.value;
                            getSiteCollections(result.value);

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

    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        $("#attachmentSelect").css("display", "block");
        var item = Office.context.mailbox.item;
        buildAttachmentsString(item.attachments);
        //$('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
        //$('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
        //$('#itemClass').text(item.itemClass);
        //$('#itemId').text(item.itemId);
        //$('#itemType').text(item.itemType);

        //$('#message-props').show();


        //$('#cc').html(buildEmailAddressesString(item.cc));
        //$('#conversationId').text(item.conversationId);
        //$('#from').html(buildEmailAddressString(item.from));
        //$('#internetMessageId').text(item.internetMessageId);
        //$('#normalizedSubject').text(item.normalizedSubject);
        //$('#sender').html(buildEmailAddressString(item.sender));
        //$('#subject').text(item.subject);
        //$('#to').html(buildEmailAddressesString(item.to));
    }

    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.

    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            //$("#attachmentSelect").append("<span>Attachments</span>");

            for (var i = 0; i < attachments.length; i++) {

                var container = $(document.createElement('div')).addClass("form-check");
                container.append('< input class="form-check-input" type ="checkbox" name="attachmentsCheck" id="attachmentCheck' + i + '" value ="' + attachments[i].name + '"/>');
                container.append('<label class="form-check-label" for="attachmentCheck' + i + '">' + attachments[i].name + '</label>');
                $("#attachmentSelect").after(container);
            }

            //$("#attachmentSelect").append('<small id="listhelp" class="form - text text - muted">Please select the attachments that needs to be saved</small>');
        } else {
            $("#attachmentSelect").append("No Attachments");
        }
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

    

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();