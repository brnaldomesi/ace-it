<!-- Original:  Tom Khoury (twaks@yahoo.com) -->
<!-- 10/22/03:  Bob Puppo - Handle hidden fields and only select if a text field -->

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
function placeFocus() {
  if (document.forms.length > 0) {
    var field = document.forms[0];
    for (i = 0; i < field.length; i++) {
      if ((field.elements[i].type !== "hidden") && ((field.elements[i].type == "text") || (field.elements[i].type == "textarea") || (field.elements[i].type == "password") || (field.elements[i].type.toString().charAt(0) == "s"))) {
        document.forms[0].elements[i].focus();
        //document.forms[0].elements[i].select();
        break;
      }
    }
  }
}
//  End -->

