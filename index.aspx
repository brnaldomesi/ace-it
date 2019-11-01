<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data"%>
<%@ Import Namespace="System.Data.SqlClient"%>
<%@ Import Namespace="System.Diagnostics"%>
<%@ Import Namespace="System.Threading"%>

<script runat="server">
    public static bool QuickOpen(SqlConnection conn, int timeout) {
        Stopwatch sw = new Stopwatch();
        bool connectSuccess = false;
        Thread t = new Thread(delegate() {
            try {
                sw.Start();
                conn.Open();
                connectSuccess = true;
            } catch { }
        });
        t.IsBackground = true;
        t.Start();
        while (timeout > sw.ElapsedMilliseconds) {
            if (t.Join(10)) {
                break;
            }
        }
        return connectSuccess;
    }
</script>

<!DOCTYPE html>
<html lang="en">
<head>
	<!-- #include file="includes/default-meta-tags.inc" -->

	<meta name="description" content="Desk review and settlement of auto and property claims, including subrogation audits. ACE was the first outsource of this type, and remains the industry leader." />
	<meta name="keywords" content="insurance, ace, ace review, desk review, claims, claim, auto, property, autoreview, subreview, subrogation, audit, audits, estimates, auto physical damage, ace-it, property, utility, propreview, demand estimate, .insurance, heavy equipment, hdt, truck, commercial, total loss, diminished value" />
    <meta name="msvalidate.01" content="B39B43D97826793073DB1D1F210E5687" />

	<title>ACE | Desk Review of Auto and Property Claims | Home</title>

	<!-- #include file="includes/jquery.inc" -->
	<!-- #include file="includes/bootstrap.inc" -->
	<!-- #include file="includes/ie-hacks.inc" -->

	<link href="css/news-flash.min.css" rel="stylesheet" />
	<link href="css/styles.css" rel="stylesheet" />
	<script src="js/email-obfuscator.js"></script>
</head>
<body>
	<div class="container-fluid">
		<!-- #include file="includes/base-content.inc" -->
		<div class="row">
		<!-- #include file="includes/topNavbar.inc" -->
		</div>
		<div class="row">
		<div class="container col-md-2" style="padding-bottom: 15px">
			<!-- #include file="includes/sidebarNav.inc" -->
			<!-- #include file="includes/newsFlash-ACE.inc" -->
		</div>

		
		<div class="container col-md-7">
			<!--- Main content - vid -->
			<div>
				<!--<style>
					#crossfade {
						position: relative;
						height: 280px;
						width: 100%;
						margin: 0 auto 10px;
					}

					#crossfade img {
						position: absolute;
						left: 0;
						-webkit-transition: opacity 1s ease-in-out;
						-moz-transition: opacity 1s ease-in-out;
						-o-transition: opacity 1s ease-in-out;
						transition: opacity 1s ease-in-out;
						opacity: 0;
						-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=0)";
						filter: alpha(opacity=0);
					}

					#crossfade img.opaque {
						opacity: 1;
						-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=100)";
						filter: alpha(opacity=1);
					}
				</style>
				<script>
					$(function () {
						String.prototype.endsWith = function (suffix) {
							return this.indexOf(suffix, this.length - suffix.length) !== -1;
						};

						var i = 0;

						(function cycleImages() {
							$("#crossfade img").removeClass("opaque");
							$("#side-menu li a").removeClass("hovered");
							$("#crossfade img").eq(i).addClass("opaque");

							if (i != 0) {
								$("#side-menu li a").eq(i - 1).addClass("hovered");
							}

							if (i == 5) {
								i = -1;
							}

							i++;
							setTimeout(cycleImages, 5000);
						})();
					});
				</script>

				<div id="crossfade">
					<img src="images/smile.png" style="width: 100%">
					<img src="images/damaged-car.jpg" style="width: 100%">
					<img src="images/home-damage.jpg" style="width: 100%">
					<img src="images/truck-repairing.jpg" style="width: 100%">
					<img src="images/car-crash.jpg" style="width: 100%;">
					<img src="images/taking_picture_of_car.png" style="width: 100%">
				</div>-->
				<div class="embed-responsive embed-responsive-16by9">
					<iframe class="embed-responsive-item" src="https://player.vimeo.com/video/263549692?autoplay=1" allowfullscreen webkitallowfullscreen mozallowfullscreen></iframe>
				</div>
				<h3 style="text-align: center">Founded in 1988, American Computer Estimating, family owned and operated, was the original outsource for desk reviews of auto physical damage claims.</h3>
				<div style="font-size: 14px;">
					ACE is the industry leader in auto, heavy equipment, subrogation, and property appraisal reviews. We are determined to remain the leader in our field by investing heavily in new technology and human resources.
				</div>
			</div>
			<div style="margin-top: 25px;">
				<!-- nope #include file="includes/newsTicker.inc" -->
				<!-- nope #include file="includes/twitterFeed.inc" -->
				<!-- nope #include file="includes/newsFlash-ACE.inc" -->
			</div>
		</div>

		<div class="container col-md-3">
			<!-- #include file="includes/login-element.inc" -->
			<!-- #include file="includes/enroll-element.inc" -->
			<!-- nope #include file="includes/newsTicker.inc" -->
			<!-- nope #include file="includes/twitterFeed.inc" -->
		</div>
		

		</div>
	</div>

	<!-- #include file="includes/footer.inc" -->
</body>
</html>
