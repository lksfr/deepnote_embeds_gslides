
/**
 * Create a open embeds menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Embed Deepnote Cell', 'showSidebar')
      .addItem('Refresh All Cells', 'refreshAll')
      .addToUi();
}

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface & adds menu items.
 */
function showSidebar() {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Embed Deepnote Cell', 'showSidebar')
      .addItem('Refresh All Cells', 'refreshAll')
      .addToUi();

  var ui = HtmlService
      .createHtmlOutputFromFile('sidebar')
      .setTitle('Embed');

  SlidesApp.getUi().showSidebar(ui);
}

/**
 * Refreshes all Deepnote embdeds in currently opened presentation.
 */
function refreshAll() {

  // check if file exists & get current presentation
  var haBDs  = DriveApp.getFilesByName("DeepNote_Slides_AddOn_MetaData_DontDelete.json")
  var presentation = SlidesApp.getActivePresentation()
  var pId = presentation.getId()

  if (!haBDs.hasNext()) {
      // Case: metadata file does not exist --> error
      throw "Invalid refresh: metadata file does not exist.";
  } else {
    // Case: metadata file does exist --> update
    var file = haBDs.next();
    var content = file.getBlob().getDataAsString();
    var json = JSON.parse(content);

    for (let slideNo = 0; slideNo < presentation.getSlides().length; slideNo++) {
      
      var allImages = presentation.getSlides()[slideNo].getImages();
      var slideImageIds = [];

      for (let i = 0; i < allImages.length; i++) {
          k = allImages[i];
          slideImageIds.push(k.getObjectId());
      }

      var matchImgIds = [];
      for (var k in json[pId]) {
        if (slideImageIds.includes(k)) {
          matchImgIds.push(k);
        }
      }

      slideImageIds = slideImageIds.sort();
      matchImgIds = matchImgIds.sort();

      for (let i = 0; i < matchImgIds.length; i++) {
        var currImgId = matchImgIds[i];
        if (slideImageIds[i]==currImgId) {
          var screenshotUrl = "https://shot.screenshotapi.net/screenshot?token=<TOKEN>" + encodeURIComponent(json[pId][currImgId]["url"]) + "&width="+json[pId][currImgId]["w"]+"&height="+json[pId][currImgId]["h"]+"&fresh=true&output=image&file_type=png&wait_for_event=load";
          allImages[i].replace(screenshotUrl)
              }
            }

    }
  }
}

/**
 * Inserts screenshot of Deepnote embed requested by user & updates metadata file on Drive.
 * @param {siteUrl} embed url string.
 * @param {w} width requested by user.
 * @param {h} height requested by user.
 * @param {slideNo} on which slide screenshot is to be inserted.
 */
function testExists(siteUrl, w, h, slideNo) {
  // build screenshot url
  var screenshotUrl = "https://shot.screenshotapi.net/screenshot?token=<TOKEN>" + encodeURIComponent(siteUrl) + "&width="+w+"&height="+h+"&fresh=true&output=image&file_type=png&wait_for_event=load";

  // get file and current presentation
  var haBDs  = DriveApp.getFilesByName("DeepNote_Slides_AddOn_MetaData_DontDelete.json")
  var presentation = SlidesApp.getActivePresentation()
  var pId = presentation.getId()

  if (!haBDs.hasNext()) {
    // Metadata file does not exist
    // create file and add first embed to metadata JSON
    presentation.getSlides()[slideNo-1].insertImage(screenshotUrl)
    var imageId = presentation.getSlides()[slideNo-1].getImages()[presentation.getSlides()[slideNo-1].getImages().length - 1].getObjectId();
    obj = {};
    obj[pId] = {};
    obj[pId][imageId] = {};
    obj[pId][imageId]["url"] = siteUrl;
    obj[pId][imageId]["w"] = w;
    obj[pId][imageId]["h"] = h;

    fileSets = {
      title: 'DeepNote_Slides_AddOn_MetaData_DontDelete.json',
      mimeType: 'application/json'
    };
  
    blob = Utilities.newBlob(JSON.stringify(obj), "application/vnd.google-apps.script+json");
    file = Drive.Files.insert(fileSets, blob);
  }
  else {
    // Metadata file does exist
    var file = haBDs.next();
    var content = file.getBlob().getDataAsString();
    var json = JSON.parse(content);

    // check if embeds for presentation already exist
    var keys = [];
    for(var k in json) keys.push(k);
    presKeyExists = keys.includes(pId);

    if (presKeyExists) {
        // Check if specific image alrEADY exists
        var embeds = [];
        for(var k in json[pId]) embeds.push(json[pId][k]);
        imgKeyExists = embeds.includes(siteUrl);

        // if embed exists -> refresh existing image
        if (imgKeyExists) {
            var matchImgId;
            for (var k in json[pId]) {
              if (json[pId][k]["url"] == siteUrl) {
                matchImgId = k;
              }
            }

            var allImages = presentation.getSlides()[slideNo-1].getImages()
            for (let i = 0; i < allImages.length; i++) {

              var currImgId = allImages[i].getObjectId();

              if (allImages[i].getObjectId()==matchImgId) {
                allImages[i].replace(screenshotUrl)
              }
            }
        } else {
        // if image doesn't exist -> insert new one & add entry to metadata & write file back
        presentation.getSlides()[slideNo-1].insertImage(screenshotUrl);
        var newImageId = presentation.getSlides()[slideNo-1].getImages()[presentation.getSlides()[slideNo-1].getImages().length - 1].getObjectId();

        json[pId][newImageId] = {};
        json[pId][newImageId]["url"] = siteUrl;
        json[pId][newImageId]["w"] = w;
        json[pId][newImageId]["h"] = h;
        
        file.setContent(JSON.stringify(json));
        }
    } else {
      // file exists but presentation has not been added yet --> insert presentation into file
        presentation.getSlides()[slideNo-1].insertImage(screenshotUrl);
        var imageId = presentation.getSlides()[slideNo-1].getImages()[presentation.getSlides()[slideNo-1].getImages().length - 1].getObjectId();

        json[pId] = {};
        json[pId][imageId] = {};
        json[pId][imageId]["url"] = siteUrl;
        json[pId][imageId]["w"] = w;
        json[pId][imageId]["h"] = h;

        file.setContent(JSON.stringify(json));
    }
  }
}

/**
 * Manages embedding request from user interface.
 * @param {siteUrl} embed url string.
 * @param {w} width requested by user.
 * @param {h} height requested by user.
 * @param {slide} on which slide screenshot is to be inserted.
 */
function embedDeepnote(siteUrl, w, h, slide) {

  // check URL
  if(siteUrl.indexOf("embed.deepnote.com")==-1) {throw "URL error. Check embed URL."}
  // call embed function
  testExists(siteUrl, w, h, Number(slide));
}