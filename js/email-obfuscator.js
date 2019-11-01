/**
 * Created by Anthony on 5/13/2015.
 */

function genMailToLink(lhs, rhs, displayText, subject) {
	document.write("<a href=\"mailto");
	document.write(":" + lhs + "%40" + rhs.toHtmlEntities());

	if (subject != null && subject.length > 0) {
		document.write("?subject=" + subject);
	}

	document.write("\">");

	if (displayText != null && displayText.length > 0) {
		document.write(displayText);
	} else {
		document.write("spam".toHtmlComment() + lhs.toHtmlEntities());
		document.write("<span style=\"display:none\">" + rand_string(Math.floor((Math.random() * 100) + 1)) + "</span>");
		document.write("@" + "bots".toHtmlComment() + rhs);
	}

	document.write("</a>");
}

/**
 * Convert a string to HTML entities.
 * @returns {string}
 */
String.prototype.toHtmlEntities = function() {
	return this.replace(/./gm, function (s) {
		return "&#" + s.charCodeAt(0) + ";";
	});
};

/**
 * Constructs an HTML comment from the string.
 * @returns {string}
 */
String.prototype.toHtmlComment = function() {
	return "<!--" + this + "-->";
};

/**
 * Creates a random string with the specified length.
 * @param length
 * @returns {string}
 */
function rand_string(length) {
	var result = "";
	var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

	for (var i = 0; i < length; i++) {
		result += possible.charAt(Math.floor(Math.random() * possible.length));
	}

	return result;
}
