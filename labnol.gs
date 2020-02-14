

function Initialize() {

  
  
  var NAME  = "Albert T. Wong";           // It will show up in the signature of your outgoing emails
  //var GROUP = "Google Script Testing";     // Enter the exact name of your Google Contacts group 
  var GROUP = "Holiday Cards";     // Enter the exact name of your Google Contacts group 




  
/*



  Tutorial: http://www.labnol.org/?p=27306
  
  Video: http://www.youtube.com/watch?v=SMxvZgK4BMg
  
  Author: Amit Agarwal (ctrlq.org)
  
  Last Updated: October 15, 2014
  
  

  H E L P
  - - - - 
     
  For help, send me a tweet @labnol or an email at amit@labnol.org




  T H E   G E E K Y   S T U F F
  - - -   - - - - -   - - - - -
  
  You can ignore everything that's below this line.  




*/
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  try {
    
    var googleGROUP = ContactsApp.getContactGroup(GROUP);
    
    if (googleGROUP) {
      
      var emailSUBJECT  = "Request for your latest contact info so we can keep in touch";    
      var myContacts = googleGROUP.getContacts();
      
      for (i=0; i<myContacts.length; i++) {
        
        var email = myContacts[i].getPrimaryEmail();
        
        if (email && email.length) {
          
          var ID = myContacts[i].getId();
          var fName = myContacts[i].getGivenName();
          ID = ID.substr(ID.lastIndexOf("/") + 1);
          
          var emailBody = "Hi " + fName + ",<br /><br />" +
            "Would you please take a moment and update your contact information in my address book. I use this information so that I can send you emails, call you and mail out holiday cards.<br /><br />" + 
              "Please <a href='" + ScriptApp.getService().getUrl() + "?id=" + 
                ID + "'>click here</a> and fill-in the required details." +
                  " Your information will be directly added to my address book." +
                    "<br /><br />" +
                   "Also so that you don't think this is an spear phishing attack to get your info, Fumie and I have 3 kids (Logan, Greyson and Rae) and at last count, we have 3 dogs, 7 cats, 60+ fishes, 2 bearded dragons, 2 snakes and 4 tortoises.  Another way to verify this email is just to email me back."   
                      
                    + "<br /><br />Thanks,<br />" + NAME + "<br /><img src='https://s3-us-west-1.amazonaws.com/public.alberttwong.com/albert_wong_signature_chop.png'><br />http://alberttwong.com - +1-949-870-9664 - GPG: 9D0F 6E75 5363 0F39 F64A 447E 2A2E 6721 C637 845A";
          
          GmailApp.sendEmail(email, emailSUBJECT, emailBody, 
                             {htmlBody:emailBody, name:NAME});
        }
      }
    }
    else
      Logger.log("Google Group not found");
  } catch (e) {
    throw e.toString();
  }
}

function doGet(e) {
  var html = HtmlService.createTemplateFromFile("form.html");
  html.id = e.parameter.id;
  var contact = labnolGetContact(e.parameter.id);
  html.home_email = contact.HOME_EMAIL;
  html.work_email = contact.WORK_EMAIL;
  html.name = contact.FULL_NAME;
  html.mobile_phone = contact.MOBILE_PHONE;
  html.home_address = contact.HOME_ADDRESS;
  //html.blog = contact.BLOG;
  html.home_page = contact.HOME_PAGE;
  return html.evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

/*
function labnolGetBasicContact(id) {    
    
  var contact = {};
  
  contact.NAME = "";
  contact.HOME_EMAIL = "";
  
  id = "http://www.google.com/m8/feeds/contacts/" + encodeURIComponent(Session.getEffectiveUser().getEmail()) + "/base/" + id; 
  
  var c = ContactsApp.getContactById(id);    
  
  if (c) {
    
    if (c.getFullName().length)
      contact.NAME = c.getFullName();
    
    if(c.getEmails(ContactsApp.Field.HOME_EMAIL).length)
      contact.HOME_EMAIL = c.getEmails(ContactsApp.Field.HOME_EMAIL)[0].getAddress();
    
  }
  Logger.log("labnolGetBasicContact" + contact.NAME + contact.HOME_EMAIL); 
  return contact;
  
}
*/

function labnolGetContact(id) {    
    
  var contact = {};  
  
  contact.FOUND = 0;
  contact.id = id;
  
  try {

    id = "http://www.google.com/m8/feeds/contacts/" + encodeURIComponent(Session.getEffectiveUser().getEmail()) + "/base/" + id; 
  
    
    var c = ContactsApp.getContactById(id);    
    
    if (c) {
      
      Logger.log(c);            
      contact.FOUND = 1;
      
      if (c.getFullName().length)
        contact.FULL_NAME = c.getFullName();
      
      if(c.getEmails(ContactsApp.Field.HOME_EMAIL).length)
        contact.HOME_EMAIL = c.getEmails(ContactsApp.Field.HOME_EMAIL)[0].getAddress();

      if(c.getEmails(ContactsApp.Field.WORK_EMAIL).length)
        contact.WORK_EMAIL = c.getEmails(ContactsApp.Field.WORK_EMAIL)[0].getAddress();      
      
      if(c.getAddresses(ContactsApp.Field.HOME_ADDRESS).length) {
        contact.HOME_ADDRESS = c.getAddresses(ContactsApp.Field.HOME_ADDRESS)[0].getAddress();
        contact.HOME_ADDRESS = contact.HOME_ADDRESS.replace(/\n/g, ", ");
      }

      if(c.getPhones(ContactsApp.Field.MOBILE_PHONE).length)
        contact.MOBILE_PHONE = c.getPhones(ContactsApp.Field.MOBILE_PHONE)[0].getPhoneNumber();
      
      //getLastUpdated()
      //getNotes()
      //getPrimaryEmail()
      
      if(c.getIMs(ContactsApp.Field.SKYPE).length)
        contact.SKYPE = c.getIMs(ContactsApp.Field.SKYPE)[0].getAddress();
            
      if(c.getUrls(ContactsApp.Field.HOME_PAGE).length)
        contact.HOME_PAGE = c.getUrls(ContactsApp.Field.HOME_PAGE)[0].getAddress();
      
      if(c.getUrls(ContactsApp.Field.BLOG).length)
        contact.BLOG = c.getUrls(ContactsApp.Field.BLOG)[0].getAddress();      
      
      if(c.getDates(ContactsApp.Field.BIRTHDAY).length) {
        var months = ["0", "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", 
                      "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
        contact.BIRTHDAY = months.indexOf(c.getDates(ContactsApp.Field.BIRTHDAY)[0].getMonth().toString()) +
          "/" + c.getDates(ContactsApp.Field.BIRTHDAY)[0].getDay() +
          "/" + c.getDates(ContactsApp.Field.BIRTHDAY)[0].getYear();
      }      
    }    
    
    return contact;
  
  } catch (e) {
  
    return contact;    
  
  }  
  
}


function labnolUpdateContact(contact) {    
  
  try {   
            
    var cid = "http://www.google.com/m8/feeds/contacts/" + encodeURIComponent(Session.getEffectiveUser().getEmail()) + "/base/" + contact.id; 

    var c = ContactsApp.getContactById(cid);   
    Logger.log(c);
    
    if (c) { 
                  
      c.setFullName(contact.FULL_NAME);

      /*
      if(c.getIMs(ContactsApp.Field.SKYPE).length)
        c.getIMs(ContactsApp.Field.SKYPE)[0].deleteIMField();      
      if (contact.SKYPE.length)
        c.addIM(ContactsApp.Field.SKYPE, contact.SKYPE);
      */
      
      if (c.getAddresses(ContactsApp.Field.HOME_ADDRESS).length)
        c.getAddresses(ContactsApp.Field.HOME_ADDRESS)[0].deleteAddressField();      
      if (contact.HOME_ADDRESS.length)
        c.addAddress(ContactsApp.Field.HOME_ADDRESS, contact.HOME_ADDRESS);
      
      if (c.getPhones(ContactsApp.Field.MOBILE_PHONE).length)
        c.getPhones(ContactsApp.Field.MOBILE_PHONE)[0].deletePhoneField();      
      if (contact.MOBILE_PHONE.length)
        c.addPhone(ContactsApp.Field.MOBILE_PHONE, contact.MOBILE_PHONE);
      
      /*
      if (c.getUrls(ContactsApp.Field.BLOG).length)
        c.getUrls(ContactsApp.Field.BLOG)[0].deleteUrlField();
      
      if (contact.BLOG.length)
        c.addUrl(ContactsApp.Field.BLOG, contact.BLOG);   
*/
      
      if (c.getUrls(ContactsApp.Field.HOME_PAGE).length)
        c.getUrls(ContactsApp.Field.HOME_PAGE)[0].deleteUrlField();      
      if (contact.HOME_PAGE.length)
        c.addUrl(ContactsApp.Field.HOME_PAGE, contact.HOME_PAGE);  

      Logger.log(contact.WORK_EMAIL);
      if (c.getEmails(ContactsApp.Field.WORK_EMAIL).length)
        c.getEmails(ContactsApp.Field.WORK_EMAIL)[0].deleteEmailField();      
      if (contact.WORK_EMAIL.length)
        c.addEmail(ContactsApp.Field.WORK_EMAIL, contact.WORK_EMAIL);  

      Logger.log(contact.HOME_EMAIL);
      if (c.getEmails(ContactsApp.Field.HOME_EMAIL).length)
        c.getEmails(ContactsApp.Field.HOME_EMAIL)[0].deleteEmailField();  
      if (contact.HOME_EMAIL.length)
        c.addEmail(ContactsApp.Field.HOME_EMAIL, contact.HOME_EMAIL);   
      
      /*
      if(contact.TWITTER.length) { 
        var cfields = c.getCustomFields();
        for (var i = 0; i < cfields.length; i++) {
          if (cfields[i].getLabel() == 'Twitter') {
            cfields[i].deleteCustomField();
          }
        }
        c.addCustomField("Twitter", "http://twitter.com/" + contact.TWITTER);
      }
      */
      if (contact.BIRTHDAY.length) {
        
        var months = 
            [ 0, ContactsApp.Month.JANUARY, ContactsApp.Month.FEBRUARY, ContactsApp.Month.MARCH,
             ContactsApp.Month.APRIL, ContactsApp.Month.MAY, ContactsApp.Month.JUNE,
             ContactsApp.Month.JULY, ContactsApp.Month.AUGUST, ContactsApp.Month.SEPTEMBER,
             ContactsApp.Month.OCTOBER, ContactsApp.Month.NOVEMBER, ContactsApp.Month.DECEMBER
            ];
        
        var date = contact.BIRTHDAY.split("/");
        
        if (c.getDates(ContactsApp.Field.BIRTHDAY).length)
          c.getDates(ContactsApp.Field.BIRTHDAY)[0].deleteDateField();
        
        c.addDate(ContactsApp.Field.BIRTHDAY, months[parseFloat(date[0])], parseFloat(date[1]), parseFloat(date[2]));
        
      }
      
      
      GmailApp.sendEmail(Session.getEffectiveUser().getEmail(),
                        "Updated: " + contact.FULL_NAME + " (" + contact.HOME_EMAIL + ")", 
                        Utilities.jsonStringify(contact));
      
      
    }    
    
  } catch (e) {
    
  }  
  
}
