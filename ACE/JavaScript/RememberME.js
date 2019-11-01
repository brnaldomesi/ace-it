// ---Below script takes into account that user may have cookies disabled so it doesn't hassle user ---//
var today = new Date();
var zero_date = new Date(0,0,0);
today.setTime(today.getTime() - zero_date.getTime());

var todays_date = new Date(today.getYear(),today.getMonth(),today.getDate(),0,0,0);
var expires_date = new Date(todays_date.getTime() + (8 * 7 * 86400000));

function storeMasterCookie() {
    if (!Get_Cookie('MasterCookie'))
        Set_Cookie('MasterCookie','MasterCookie');
}

function storeIntelligentCookie(name,value) {
    if (Get_Cookie('MasterCookie')) {
        var IntelligentCookie = Get_Cookie(name);
        if ((!IntelligentCookie) || (IntelligentCookie != value)) {
            Set_Cookie(name,value,expires_date);
            var IntelligentCookie = Get_Cookie(name);
            if ((!IntelligentCookie) || (IntelligentCookie != value))
                Delete_Cookie('MasterCookie');
        }
    }
}

function FixCookieDate (date) {
  // Function to correct for 2.x Mac date bug.  Call this function to
  //  fix a date object prior to passing it to SetCookie.
  //  IMPORTANT:  This function should only be called *once* for
  //  any given date object!  See example at the end of this document.
  //

  var base = new Date(0);
  var skew = base.getTime(); // dawn of (Unix) time - should be 0

  if (skew > 0)  // Except on the Mac - ahead of its time :-)
    date.setTime (date.getTime() - skew);
}


// --- Cookie Handling -- //
function Get_Cookie(name) {
    var start = document.cookie.indexOf(name+"=");
    var len = start+name.length+1;
    if ((!start) && (name != document.cookie.substring(0,name.length))) return null;
    if (start == -1) return null;
    var end = document.cookie.indexOf(";",len);
    if (end == -1) end = document.cookie.length;
    return unescape(document.cookie.substring(len,end));
}

function Set_Cookie(name,value,expires,path,domain,secure) {
  /*
     name - name of the cookie
     value - value of the cookie
     [expires] - expiration date of the cookie (defaults to end of current session)
     [path] - path for which the cookie is valid (defaults to path of calling document)
     [domain] - domain for which the cookie is valid (defaults to domain of calling document)
     [secure] - Boolean value indicating if the cookie transmission requires a secure transmission
     * an argument defaults when it is assigned null as a placeholder
     * a null placeholder is not required for trailing omitted arguments
  */
    document.cookie = name + "=" +escape(value) +
        ( (expires) ? ";expires=" + expires.toGMTString() : "") +
        ( (path) ? ";path=" + path : "") + 
        ( (domain) ? ";domain=" + domain : "") +
        ( (secure) ? ";secure" : "");
}

function Delete_Cookie(name,path,domain) {
  /*
     name - name of the cookie
     [path] - path of the cookie (must be same as path used to create cookie)
     [domain] - domain of the cookie (must be same as domain used to create cookie)
     path and domain default if assigned null or omitted if no explicit argument proceeds
  */

    if (Get_Cookie(name)) document.cookie = name + "=" +
       ( (path) ? ";path=" + path : "") +
       ( (domain) ? ";domain=" + domain : "") +
       ";expires=Thu, 01-Jan-70 00:00:01 GMT";
}

// --- Remember me helper functions --- //
function RememberMe (f) {
    var ExpireDate = new Date();
    //ExpireDate = FixCookieDate(ExpireDate);  // don't need since time doesn't need to be exact
	ExpireDate.setTime (ExpireDate.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 days from now 
    
    Set_Cookie('txtUserName', f.txtUserName.value, ExpireDate, '/', '.staging-portal.ace-it.com', '');
    Set_Cookie('txtPassword', f.txtPassword.value, ExpireDate, '/', '.staging-portal.ace-it.com', '');
    Set_Cookie('chkRememberMe', f.RememberMe.value, ExpireDate, '/', '.staging-portal.ace-it.com', '');
}

function ForgetMe (f) {
    Delete_Cookie('txtUserName', '/', '.staging-portal.ace-it.com');
    Delete_Cookie('txtPassword', '/', '.staging-portal.ace-it.com');
    Delete_Cookie('chkRememberMe', '/', '.staging-portal.ace-it.com');
    f.txtUserName.value = '';
    f.txtPassword.value = '';
    f.RememberMe.value = '';
}

function RetrievMe (f) {
  if (Get_Cookie("chkRememberMe") == "1") 
    f.chkRememberMe.checked = true;

  if (Get_Cookie("chkRememberMe") == "1") {
    if (Get_Cookie("txtUserName")) 
      f.txtUserName.value = Get_Cookie("txtUserName");
    if (Get_Cookie("txtPassword")) 
      f.txtPassword.value = Get_Cookie("txtPassword");  
  }
}
