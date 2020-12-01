const ENTERPRISE_PRODUCT = '101031'; //The product ID for Google Enterprise for Education
const TEACHER_LICENSE = '1010310002';// The SKU ID for Teacher License
const STUDENT_LICENSE = '1010310003';// The SKU ID for Student License
const DOMAIN = YOUR_DOMAIN

function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Licenses')
        .addItem('Show licenses', 'showLicenses')
        .addItem('Add and Remove licenses', 'addAndRemoveLicenses')
        .addToUi();    
}

function showLicenses() {
  schools("read");
}

function addAndRemoveLicenses() {
  schools("write");
}

function schools(rw) {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet();
  //Find the list of school names in the first column of the Sheet "Schools" or "Schools not ready"
  var schoolSheet = (rw == "write") ? mySheet.getSheetByName("Schools"): mySheet.getSheetByName("Schools not ready") ;
  var numRows = schoolSheet.getLastRow();
  var schoolRange = schoolSheet.getRange(1,1,numRows,1);
  schoolRange.getValues().forEach(function (school) {
    var schoolname = ""+school;
    if( mySheet.getSheetByName(schoolname) != null ) {
      mySheet.deleteSheet(mySheet.getSheetByName(schoolname));
      mySheet.insertSheet(schoolname);
    } else {
      mySheet.insertSheet(schoolname);
    }
    showUsers(schoolname, rw);
  });

}

function showUsers(school, rw) {
  //Fill a sheet for each school
  var schoolSheet = SpreadsheetApp.getActive().getSheetByName(school);
  var pageToken;
  var page;
  var productId = ENTERPRISE_PRODUCT;
  var skuId = '';
  //Get users from OU Staff of Each School
  var query = "orgUnitPath='/"+school+"/Staff'";
  do {
    page = AdminDirectory.Users.list({
      domain: DOMAIN,
      orderBy: 'givenName',
      maxResults: 100,
      pageToken: pageToken,
      query: query
    });
    var users = page.users;
    if (users) {
      users.forEach(function (user) {
        var hasLicense = false;
        var description = "";
        try {
          description = user.organizations[0].description;
        } catch (e) {
          description = "None";
        }                
        var title = "";
        try {
          title = user.organizations[0].title;
        } catch (e) {
          title = "None";
        }
        
        //If the user description contains "Lærer", "Leder" or "Ledelse", they shoud have license.
        var shouldHaveLicense = ( !user.suspended && (description.includes("Lærer") || description.includes("Leder") || description.includes("Ledelse")));
        //If the user description contains "Vikar", assign student-license. TODO: Praktikant
        skuId = description.includes("Vikar") ? STUDENT_LICENSE : TEACHER_LICENSE;
        //Find license settings for every user in the OU
        var license = ""; 
        try {
          license = AdminLicenseManager.LicenseAssignments.get(productId, skuId, user.primaryEmail).skuName;
          hasLicense = true;
        } catch(e) {
          license = "No license";
          hasLicense = false;
        } 

        //Assign license to eligible users or remove licenses to users, that are not.
        var results = "";
        if( rw.includes("write") ) {
          var userId = user.primaryEmail;
          if( shouldHaveLicense && !hasLicense ) {
            try {
              results = AdminLicenseManager.LicenseAssignments.insert({userId: userId}, productId, skuId);
            } catch(e) {
              try {
                results = AdminLicenseManager.LicenseAssignments.update({productId: productId, skuId: skuId, userId: userId}, productId, TEACHER_LICENSE, userId);
              } catch (err) {
                results = err + " " + e;
              }
            }
          } else if( (shouldHaveLicense && hasLicense)  ) {
            results = "Ok - Has license";
          } else if( !shouldHaveLicense && !hasLicense ) {
            results = "Ok - Need no license";
          } else {
            try { 
              AdminLicenseManager.LicenseAssignments.remove(productId, skuId, userId);
              results = "License removed";
            } catch(e) {
              results = e;
            }
          }
        }
        schoolSheet.appendRow([user.name.fullName,user.primaryEmail,title,description,"Need license:",shouldHaveLicense,"Suspended:",user.suspended,"License",license,"Last activity", user.lastLoginTime,"Result:",results]);
        if( !shouldHaveLicense && (user.lastLoginTime.includes("2020"))){
          SpreadsheetApp.getActiveSheet().getRange(schoolSheet.getLastRow(),1,1,SpreadsheetApp.getActiveSheet().getLastColumn()).setBackgroundRGB(224, 102, 102);
        }
      });
    } else {
      Logger.log('No users found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}
