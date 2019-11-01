How can I mimic a client-side POST from ASP?     (3,267 requests)
Often you need to POST information to an ASP or other script without relying on the user to click a Submit button. There are a few ways to work around this: 
 


Write a component. We have one that does this (among many other things); however, it is proprietary, and as such I can't distribute source. I will tell you it is an ATL component and makes a low-level connection.
Use an existing HTTP component, such as AspHTTP
"Fake" a POST operation with client-side script, as in the following: 
 
    <form 
        method=post 
        action='<script>' 
        name='myform'> 
 
    <input 
        type=hidden 
        name='myname' 
        value='myvalue'> 
 
    </form> 
 
    <script> 
        window.onLoad = document.myform.submit(); 
    </script> 
 
 
 
Note that it's fairly easy to build such a form with ASP, simply by iterating through the incoming form elements 
 

