// commented <script> tag here makes interdev colorize properly
//
//  jsrsClient.asp - javascript remote scripting client include
//  
//  Author:  Brent Ashley [jsrs@megahuge.com]
//
//  make asynchronous remote calls to server without client page refresh
//
//  see license.txt for copyright and license information

/*

ver   date       comment
-----------------------
1.00  22Aug2000  release
1.01  06Sep2000  change location.href=url to location.replace(url) to help avoid history
1.02  12Sep2000  comment out unique string to reduce history clutter
1.03  22Sep2000  add jsrsEvalEscape so quotes and newlines in parms won't choke
1.04  23Jan2001  ie4.x bug fix, debug utility: 
  hidden iframes must be inserted in space which has been visibly rendered on screen.  
  - moved insertion point to top of body to reduce/eliminate likelihood of this being unrendered space.
  - changed DIV to SPAN to avoid line break insertion into page flow
  added jsrsDebugInfo()
  - see comment for instruction on how to use it
1.05  06Feb2001  escape jsrsError return value, use try/catch in dispatch
1.06  13Feb2001  added no-nonsense copyright and license, starting to plan for NS6

1.1   14Feb2001  Mozilla/NS6 compatibility
  - the code expects NS4, IE4+, or Mozilla (incl NS6).
  - it's up to you to tell the user if they don't have the right browser.
1.11  15Feb2001  NS6 was not reloading on recall 
  - set src to '' before reloading
  - thanks to Murray Cash

*/

// callback pool needs global scope
var jsrsContextPoolSize = 0;
var jsrsContextMaxPool = 10;
var jsrsContextPool = new Array();
var jsrsBrowser = jsrsBrowserSniff();

// constructor for context object
function jsrsContextObj( contextID ){
  
  // properties
  this.id = contextID;
  this.busy = true;
  this.callback = null;
  this.container = contextCreateContainer( contextID );
  
  // methods
  this.callURL = contextCallURL;
  this.getPayload = contextGetPayload;
  this.setVisibility = contextSetVisibility;
}

//  method functions are not privately scoped 
//  because Netscape's debugger chokes on private functions
function contextCreateContainer( containerName ){
  // creates hidden container to receive server data 
  var container;
  switch( jsrsBrowser ) {
    case 'NS':
      container = new Layer(100);
      container.name = containerName;
      container.visibility = 'hidden';
      container.clip.width = 100;
      container.clip.height = 100;
      break;
    
    case 'IE':
      document.body.insertAdjacentHTML( "afterBegin", '<span id="SPAN' + containerName + '"></span>' );
      var span = document.all( "SPAN" + containerName );
      var html = '<iframe name="' + containerName + '" src=""></iframe>';
      span.innerHTML = html;
      span.style.display = 'none';
      container = window.frames[ containerName ];
      break;
      
    case 'MOZ':  
      var span = document.createElement('SPAN');
      span.id = "SPAN" + containerName;
      document.body.appendChild( span );
      var iframe = document.createElement('IFRAME');
      iframe.name = containerName;
      span.appendChild( iframe );
      container = iframe;
      break;
  }
  return container;
}

function contextCallURL( URL ){
  switch( jsrsBrowser ) {
    case 'NS':
      this.container.src = URL;
      break;
    case 'IE':
      this.container.document.location.replace(URL);
      break;
    case 'MOZ':
      this.container.src = '';
      this.container.src = URL; 
      break;
  }  
}

function contextGetPayload(){
  switch( jsrsBrowser ) {
    case 'NS':
      return this.container.document.forms['jsrs_Form'].elements['jsrs_Payload'].value;
    case 'IE':
      return this.container.document.forms['jsrs_Form']['jsrs_Payload'].value;
    case 'MOZ':
      return window.frames[this.container.name].document.forms['jsrs_Form']['jsrs_Payload'].value; 
  }  
}

function contextSetVisibility( vis ){
  switch( jsrsBrowser ) {
    case 'NS':
      this.container.visibility = (vis)? 'show' : 'hidden';
      break;
    case 'IE':
      document.all("SPAN" + this.id ).style.display = (vis)? '' : 'none';
      break;
    case 'MOZ':
      document.getElementById("SPAN" + this.id).style.visibility = (vis)? '' : 'hidden';
      this.container.width = (vis)? 250 : 0;
      this.container.height = (vis)? 100 : 0;
      break;
  }  
}

// end of context constructor

function jsrsGetContextID(){
  var contextObj;
  for (var i = 1; i <= jsrsContextPoolSize; i++){
    contextObj = jsrsContextPool[ 'jsrs' + i ];
    if ( !contextObj.busy ){
      contextObj.busy = true;      
      return contextObj.id;
    }
  }
  // if we got here, there are no existing free contexts
  if ( jsrsContextPoolSize <= jsrsContextMaxPool ){
    // create new context
    var contextID = "jsrs" + (++jsrsContextPoolSize);
    jsrsContextPool[ contextID ] = new jsrsContextObj( contextID );
    return contextID;
  } else {
    alert( "jsrs Error:  context pool full" );
    return null;
  }
}

function jsrsExecute( rspage, callback, func, parms, visibility ){
  // call a server routine from client code
  //
  // rspage      - href to asp file
  // callback    - function to call on return 
  //               or null if no return needed
  //               (passes returned string to callback)
  // func        - sub or function name  to call
  // parm        - string parameter to function
  //               or array of string parameters if more than one
  // visibility  - optional boolean to make container visible for debugging
  
  // get context
  var contextObj = jsrsContextPool[ jsrsGetContextID() ];
  contextObj.callback = callback;

  var vis = (visibility == null)? false : visibility;
  contextObj.setVisibility( vis );

  // build URL to call
  var URL = rspage;

  // always send context
  URL += "?C=" + contextObj.id;
  
  // func and parms are optional
  if (func != null){
    URL += "&F=" + escape(func);

    if (parms != null){
      if (typeof(parms) == "string"){
        // single parameter
        URL += "&P0=[" + escape(parms+'') + "]";
      } else {
        // assume parms is array of strings
        for( var i=0; i < parms.length; i++ ){
          URL += "&P" + i + "=[" + escape(parms[i]+'') + "]";
        }
      } // parm type
    } // parms
  } // func

  // make the call
  contextObj.callURL( URL );
  
  return contextObj.id;
}

function jsrsLoaded( contextID ){
  // get context object and invoke callback
  var contextObj = jsrsContextPool[ contextID ];
  if( contextObj.callback != null){
    contextObj.callback( jsrsUnescape( contextObj.getPayload() ), contextID );
  }
  // clean up and return context to pool
  contextObj.callback = null;
  contextObj.busy = false;
}

function jsrsError( contextID, str ){
  alert( unescape(str) );
  jsrsContextPool[ contextID ].busy = false
}

function jsrsUnescape( str ){
  // payload has slashes escaped with whacks
  return str.replace( /\\\//g, "/" );
}

function jsrsBrowserSniff(){
  if (document.layers) return "NS";
  if (document.all) return "IE";
  if (document.getElementById) return "MOZ";
  return "OTHER";
}

/////////////////////////////////////////////////
//
// user functions

function jsrsArrayFromString( s, delim ){
  // rebuild an array returned from server as string
  // optional delimiter defaults to ~
  var d = (delim == null)? '~' : delim;
  return s.split(d);
}

function jsrsDebugInfo(){
  // use for debugging by attaching to f1 (works with IE)
  // with onHelp = "return jsrsDebugInfo();" in the body tag
  var doc = window.open().document;
  doc.open;
  doc.write( 'Pool Size: ' + jsrsContextPoolSize + '<br><font face="arial" size="2"><b>' );
  for( var i in jsrsContextPool ){
    var contextObj = jsrsContextPool[i];
    doc.write( '<hr>' + contextObj.id + ' : ' + (contextObj.busy ? 'busy' : 'available') + '<br>');
    doc.write( contextObj.container.document.location.pathname + '<br>');
    doc.write( contextObj.container.document.location.search + '<br>');
    doc.write( '<table border="1"><tr><td>' + contextObj.container.document.body.innerHTML + '</td></tr></table>' );
  }
  doc.write('</table>');
  doc.close();
  return false;
}
