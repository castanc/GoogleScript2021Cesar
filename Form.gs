


function onOpen() {
  Logger.log("onOpen()");
  //activateFormSubmit();


}

//todo: if the form is modified the trigger must be recreated
function activateFormSubmit() {
  Logger.log("activateSubmit()");

  let form = FormApp.getActiveForm();
  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();
  Logger.log("ActivateSubmit() compelted");

}

//https://support.google.com/docs/thread/27649980/script-to-get-information-from-google-form-create-copy-of-a-spreadsheet-and-give-permissions?hl=en


function onFormSubmit(e) {
  Logger.log("on form submit");
  Logger.log(JSON.stringify(e));


  Logger.log(`email; ${e.response.getRespondentEmail()}`);

  Logger.log("authMode=%s, source.getId()=%s", e.authMode, e.source.getId());
  var items = e.response.getItemResponses();

  let namedValues = {}
  let info = "";
  for (i in items) {
    Logger.log(`Title:${items[i].getItem().getTitle()} Value:[${items[i].getResponse()}]`);

    let item = items[i].getItem();
    Logger.log(`i:${i} items[i]: ${items[i]} item: ${JSON.stringify(item)}`);
    if (items[i].getResponse().toString().lenght > 0)
    {
      info = `${info}${items[i].getItem().getTitle()}|${items[i].getResponse()};`;
      namedValues.push({"Name":items[i].getItem().getTitle(),
      "Value:":items[i].getResponse()});
    }
  }
  Logger.log(info);
}




