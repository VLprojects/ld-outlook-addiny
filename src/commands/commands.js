/**
 * Ensures the Office.js library is loaded.
 */
const requestUrl = 'https://moodhood-api.livedigital.space/v1/';
let headers = {
  "Access-Control-Allow-Origin": "*",
  "Content-Type": "application/json",
  "access-control-allow-credentials" : "true" ,
  "vary": "Origin"
};

Office.onReady((info) => {
  /** 
   * Maps the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
   * This ensures support in Outlook on Windows. 
   */

  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);

});

function onAppointmentSendHandler(event) {

  let _roomid = localStorage.getItem("roomid");
  let _spaceid =localStorage.getItem("spaceid")
  let token = localStorage.getItem("token");

  
  if (!token) {
    event.completed({
      allowEvent: false,
      errorMessage: "Нужно авторизовать в плагине для сохранения мероприятия ",
      cancelLabel: "Login",
      sendAnywayLabel: "Отправить",
      commandId: "msgReadOpenPaneButton"
    });
    return;
  }
  let _event_info = {
    "name": `002 | `,
    "isPublic": true,
    "isScreensharingAllowed": true,
    "isChatAllowed": true,
    "type": "lesson"
  };

  Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, (asyncResult) => {
    const event = asyncResult.asyncContext;
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    }

    let _subject = asyncResult.value;
    Office.context.mailbox.item.body.getAsync("html",{ asyncContext: event }, (asyncResult) => {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
      let _body = asyncResult.value;
      console.log("_body =", _body);
      
      var _start = _body.indexOf('<a href=');
      var _end = _body.indexOf('</a>')
      _body = _body.slice(0, _start) + _subject + _body.slice(_end + 4);

      Office.context.mailbox.item.body.setAsync(_body,
        { coercionType: "html", asyncContext: event},
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }

            // save room when send
            _event_info["name"] = _subject;
            try {
              $.ajax({
                url: requestUrl+'spaces/'+_spaceid+'/rooms/'+_roomid,
                method: "PUT",
                cors: true ,
                secure: true,                
                headers: {
                  ...headers,
                  "Authorization": 'Bearer '+token,
                },
                data:JSON.stringify(_event_info)
              }).done(function(response){
                // callback(gists);
                console.log(response);

                // w/out prompt
                asyncResult.asyncContext.completed({ allowEvent: true });
                localStorage.setItem("tmp_to_delete", "");


              }).fail(function(error){
                console.log("err",error);
                asyncResult.asyncContext.completed({
                  allowEvent: false,
                  errorMessage: "Ошибка сервера ",
                  cancelLabel: "Отменить",
                  commandId: "msgReadOpenPaneButton"
                });       
              });
            } catch (error) {
                asyncResult.asyncContext.completed({
                  allowEvent: false,
                  errorMessage: "Ошибка сети ",
                  cancelLabel: "отменить",
                  commandId: "msgReadOpenPaneButton"
                }); 
                
            }
          });
        }
      );
    });

}

function onNewAppointmentComposeHandler(event) {
  const newRecipients = [
      {
          "displayName": Office.context.mailbox.userProfile.displayName,
          "emailAddress": Office.context.mailbox.userProfile.emailAddress
      }
  ];
  Office.context.mailbox.item.requiredAttendees.setAsync(newRecipients, {
    "asyncContext": event
  },
  function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`Successfully added _alias ${JSON.stringify(asyncResult.status)}`);
    } else {
      console.error(`Failed to add locations. Error message: ${asyncResult.error.message}`);
    }
    });
  let _spaceid = localStorage.getItem("spaceid");
  let _deleteList = localStorage.getItem("tmp_to_delete");
    
  deleteDiscards();
}

function deleteDiscards() {
  console.log("deleteDiscards() started");
  let token = localStorage.getItem("token");
  // localStorage.setItem("tmp_to_delete", "");
  
  let _deleteList = localStorage.getItem("tmp_to_delete");
  console.log("tmp_to_delete _deleteList = ",_deleteList);
  if (_deleteList) {
    _deleteList = JSON.parse(_deleteList);
    let array = _deleteList;
    for (let i = 0; i < array.length; i++) {
      $.ajax({
        url: requestUrl + 'spaces/' + array[i].spaceId + '/rooms/' + array[i].roomId,
        method: "DELETE",
        cors: true,
        secure: true,
        headers: {
          ...headers,
          "Authorization": 'Bearer ' + token,
        },
      }).done(function (response) {
        console.log("deleteDiscards() .done");
        console.log(response);
        localStorage.setItem("tmp_to_delete", "");
        
      }).fail(function (error) {
        console.log("deleteDiscards() .fail  error =", error);
      });
    }
  }
}
      



