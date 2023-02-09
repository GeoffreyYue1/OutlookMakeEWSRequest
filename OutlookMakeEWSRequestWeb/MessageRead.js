(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
        loadProps();

       
    });

      $('#btnReportPhishing').click(btnReportPhishingClick);

      $('#btnDeleteMail').click(btnDeleteMailClick);

    };

    function btnDeleteMailClick() {
        const mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(MoveItemToDeletion(mailbox.item.itemId), function (result) {
            const response = $.parseXML(result.value);
            console.log(response);

        });
    }

    function MoveItemToDeletion(itemId) {
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header>' +
            '<t:RequestServerVersion Version="Exchange2013" />' +
            '</soap:Header>' +
            '<soap:Body>' +
            '<m:MoveItem>' +
            '<m:ToFolderId>' +
            '<t:DistinguishedFolderId Id="recoverableitemsdeletions" />' +
            '</m:ToFolderId>' +
            '<m:ItemIds>' +
            //'<t:ItemId Id="AAMkADgxZmJlZTZkLTE5NzQtNDNiMC1hMGNmLTdiMzM3ODdjODI1OABGAAAAAADpJU+q18HkQ5SY/3kcz4IABwABmZcuHwycTL8Q30ZGi/DwAAAAAAEMAAABmZcuHwycTL8Q30ZGi/DwAAJEBdSFAAA=" ChangeKey="CQAAABYAAAABmZcuHwycTL8Q30ZGi/DwAAJEBIKp" />' +
            '<t:ItemId Id="' + itemId + '" />' +
            '</m:ItemIds>' +
            '</m:MoveItem>' +
            '</soap:Body>' +
            '</soap:Envelope >';

        console.log(request);
        return request;
}

    function btnReportPhishingClick() {

        const mailbox = Office.context.mailbox;
        mailbox.makeEwsRequestAsync(getItemRequest(mailbox.item.itemId), function (result) {
            const response = $.parseXML(result.value);
            console.log(response);


            const mimecontent = response.getElementsByTagName("t:MimeContent")[0];
            console.log(mimecontent);
            const stringmime = mimecontent.textContent;
            console.log(stringmime);


            const tsubject = response.getElementsByTagName("t:Subject")[0];
            console.log(tsubject);
            const stringSubject = tsubject.textContent;
            console.log(stringSubject);

            mailbox.makeEwsRequestAsync(CreateItemWithMIMEContentRequest("test002", "This is a test message", "u8@geoffrey1.onmicrosoft.com",stringSubject,stringmime), function (asyncResult0) {

                const CreateItemResponse = $.parseXML(asyncResult0.value);
                console.log(CreateItemResponse);

                const NewItemId = CreateItemResponse.getElementsByTagName("t:ItemId")[0];
                console.log(NewItemId);

                const strNewItemId = NewItemId.attributes["Id"].value;
                console.log(strNewItemId);
                const strNewItemChangeKey = NewItemId.attributes["ChangeKey"].value;
                console.log(strNewItemChangeKey);

                mailbox.makeEwsRequestAsync(SendItemRequest(strNewItemId, strNewItemChangeKey), function (asyncResult1) {
                    const SendItemResponse = $.parseXML(asyncResult1.value);
                    console.log(SendItemResponse);
                });


                });
            

        });

    }

    
    

    function CreatItemcallback(asyncResult) {
       

    }

    function AddAttachmentcallback(asyncResutl) {
       
    }

    function SendItemRequest(draftitemid, draftchangekey) {
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header>' +
            '<t:RequestServerVersion Version="Exchange2013" />' +
            '</soap:Header>' +
            '<soap:Body>' +
            '<m:SendItem SaveItemToFolder="true">' +
            '<m:ItemIds>' +
            '<t:ItemId Id="' + draftitemid + '" ChangeKey="' + draftchangekey + '" />' +
            '</m:ItemIds>' +
            '<m:SavedItemFolderId>' +
            '<t:DistinguishedFolderId Id="sentitems" />' +
            '</m:SavedItemFolderId>' +
            '</m:SendItem>' +
            '</soap:Body>' +
            '</soap:Envelope>';
        console.log(request);
        return request;
    }

    function CreateItemWithMIMEContentRequest(subject,body,toAddress,attachemntSubject, attachmentMIMEContent) {

       
        // Return a GetItem operation request for the subject of the specified item.
        const request =
            '<?xml version="1.0" encoding="utf-8" ?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
            '<soap:Header>' +
            '<t:RequestServerVersion Version="Exchange2013" />' +
            '</soap:Header>' +
            '<soap:Body>' +
            '<m:CreateItem MessageDisposition="SaveOnly">' +
            '<m:Items>' +
            '<t:Message>' +
            '<t:Subject>' + subject
            + '</t:Subject>' +
            '<t:Body BodyType="HTML">' +
            body

            + '</t:Body>' +
            '<t:Attachments>' +
            '<t:ItemAttachment>' +
            '<t:Name>' +
            attachemntSubject
            + '</t:Name>' +
            '<t:IsInline>false</t:IsInline>' +
            '<t:Message>' +
            '<t:MimeContent CharacterSet="UTF-8">' +
            attachmentMIMEContent
            + '</t:MimeContent>' +
            '</t:Message>' +
            '</t:ItemAttachment>' +
            '</t:Attachments>' +
            '<t:ToRecipients>' +
            '<t:Mailbox>' +
            '<t:EmailAddress>' +
            toAddress
            + '</t:EmailAddress>' +
            '</t:Mailbox>' +
            '</t:ToRecipients>' +
            '</t:Message>' +
            '</m:Items>' +
            '</m:CreateItem>' +
            '</soap:Body>' +
            '</soap:Envelope>';
        console.log(request);
        return request;
    }

    function getItemRequest(itemId) {

        console.log(itemId);
        // Return a GetItem operation request for the subject of the specified item.
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
            '      <ItemShape>' +
            '        <t:BaseShape>IdOnly</t:BaseShape>' +
            '        <t:AdditionalProperties>' +
            '            <t:FieldURI FieldURI="item:Subject"/>' +
            '            <t:FieldURI FieldURI="item:MimeContent"/>' +
            '        </t:AdditionalProperties>' +
            '      </ItemShape>' +
            '      <ItemIds><t:ItemId Id="' + itemId + '"/></ItemIds>' +
            '    </GetItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';
        return request;
    }

     function createItemRequest(subject) {
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '<soap:Body>' +
            '<m:CreateItem MessageDisposition="SaveOnly">' +
            '<m:Items>' +
            '<t:Message>' +
            '<t:Subject>' + subject + '</t:Subject>' +
            '<t:Body BodyType="HTML">{ submittedBy: ' + Office.context.mailbox.userProfile.emailAddress +
            // 'reason: '+reason+'<br/>'
            // 'opened: '+opened+'<br/>'
            '</t:Body>' +
            '<t:ToRecipients>' +
            '<t:Mailbox>' +
            '<t:EmailAddress>' + Office.context.mailbox.userProfile.emailAddress + '</t:EmailAddress>' +
            '</t:Mailbox>' +
            '</t:ToRecipients>' +
            '</t:Message>' +
            '</m:Items>' +
            '</m:CreateItem>' +
            '</soap:Body>' +
            '</soap:Envelope>';

        return request;
    }


    function addAttachment(itemId, mimecontent, changeKey) {
        console.log("Add attachment");
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013"/>' +
            '  </soap:Header>' +
            '<soap:Body>' +
            '<m:CreateAttachment>' +
            '<m:ParentItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /> ' +
            '<m:Attachments>' +
            '<t:ItemAttachment>' +
            '<t:Name>' + Office.context.mailbox.item.subject + '</t:Name>' +
            '<t:IsInline>false</t:IsInline>' +
            '<t:Message>' +
            '<t:MimeContent CharacterSet="UTF-8">' + mimecontent + '</t:MimeContent>' +
            '</t:Message>' +
            '</t:ItemAttachment>' +
            '</m:Attachments>' +
            '</m:CreateAttachment>' +
            '</soap:Body>' +
             '</soap:Envelope>';

         console.log(request);

        return request;
    }

     function addAttachment2(itemId, mimecontent, changeKey) {
        const request =
            '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +


            '<soap:Body>' +
            '<CreateAttachment xmlns="https://schemas.microsoft.com/exchange/services/2006/messages"' +
            '                  xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types">' +
            '<ParentItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /> ' +
            '<Attachments>' +
            '<t:ItemAttachment>' +
            '<t:Name>' + Office.context.mailbox.item.subject + '</t:Name>' +
            '<t:IsInline>false</t:IsInline>' +
            '<t:Message>' +
            '<t:MimeContent CharacterSet="UTF-8">' + mimecontent + '</t:MimeContent>' +
            '</t:Message>' +
            '</t:ItemAttachment>' +
            '</Attachments>' +
            '</CreateAttachment>' +
            '</soap:Body>' +
            '</soap:Envelope>';
        return request;
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