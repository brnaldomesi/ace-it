<style>
	ul.socialmedia-list {
		list-style-type: none;
		float: right;
		vertical-align: middle;
	}

	ul.socialmedia-list li {
		display: inline-block;
		vertical-align: middle;
		margin-top: 10px;
	}

	ul.socialmedia-list li a {
		display: block;
		background: url('images/socialmedia_sprites.png') no-repeat;
		height: 24px;
		width: 24px;
		color: transparent;
	}

	.facebook {
		background-position: 0 0 !important;
	}

	.facebook:hover {
		background-position: 0 -24px !important;
	}

	.twitter {
		background-position: -24px 0 !important;
	}

	.twitter:hover {
		background-position: -24px -24px !important;
	}

	.linkedin {
		background-position: -48px 0 !important;
	}

	.linkedin:hover {
		background-position: -48px -24px !important;
	}

	.side-menu {
		/* insert old font maybe */
		font-style: italic;
		font-size: large;
	}

	.hovered {
		background-color: lightgray;
	}
</style>

<script>
	function killCopy(e) {
		return false;
	}
		
	function reEnable() {
		return true;
	}
	
	$(function() {
		var titleElements = document.title.split("|");
		var lastElement = titleElements[titleElements.length - 1].trim().replace(" ", "-").toLowerCase();

		if (lastElement != "home" && lastElement != "enroll") { // firefox 40 screws up forms
			document.onselectstart = new Function("return false");
		
			if (window.sidebar) {
				document.onmousedown = killCopy;
				document.onclick = reEnable;
			}
		}
	});
</script>

<table style="border-collapse: collapse; width:100%">
	<tbody>
		<tr>
			<td style="text-align: center;">
				<a href="index.aspx"><img src="images/ace_logo_with_motto.gif" style="height: 10em; padding: 5px"></a>
			</td>
			<td class="text-right; width:100%">
				<img src="images/ace-header.png" width="100%" height="auto">
			</td>
		</tr>
	</tbody>
</table>

<script>
	function appendClassByID(id, otherClasses) {
		var element = document.getElementById(id);

		if (element != null) {
			element.className = (element.className != null) ? element.className + " " + otherClasses : otherClasses;
		}
	}

	$(function() {
		var titleElements = document.title.split("|");
		var lastElement = titleElements[titleElements.length - 1].trim().replace(" ", "-").toLowerCase();

		appendClassByID(lastElement + "-tab", "active");
		$("#" + lastElement).addClass("hovered");
	});
</script>
