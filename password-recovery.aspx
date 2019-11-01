<!DOCTYPE html>
<html lang="en">
<head>
	<!-- #include file="includes/default-meta-tags.inc" -->

	<title>ACE | Desk Review of Auto and Property Claims | Password Recovery</title>

	<!-- #include file="includes/jquery.inc" -->
	<!-- #include file="includes/bootstrap.inc" -->
	<!-- #include file="includes/ie-hacks.inc" -->

	<link href="css/styles.css" rel="stylesheet" />
</head>
<body>
	<div class="container">
		<!-- #include file="includes/base-content.inc" -->
		<div class="row">
		<!-- #include file="includes/topNavbar.inc" -->
		</div>
		<div class="row">
		<div class="container col-md-2" style="padding-bottom: 15px">
			<!-- #include file="includes/sidebarNav.inc" -->
		</div>

		<div class='col-md-10'>
        <div class="well">
            <form name="frmForgotPass" method="post" action="/ACE/ForgotPass.asp?txtState=1" class="form-horizontal">
                <fieldset>
                    <legend>
                        Password Recovery
                    </legend>
                    <div class="row text-center">
                        <h4>User Information</h4>
                    </div>
                    <br/>
                    <div class="form-group">
                        <label for="txtUserName" class="col-md-2 control-label">User ID</label>
                        <div class="col-md-10">
                            <input type="text" class="form-control" name="txtUserName" id="txtUserName" autofocus required>
                        </div>
                    </div>
					<div class="form-group">
                        <div class="col-md-10 col-md-offset-2">
                            <button type="submit" name="submitup" id="submitup" class="btn btn-primary btn-md">Submit</button>
                        </div>
                    </div>
                    <br/>
                    <div class="form-group text-center">
                        <div class="col-md-10 col-md-offset-2">
                            <b>
                                A message containing your password will be sent to the email address that you provided when you created your staging-portal.ace-it.com account.
                            </b>
                        </div>
                    </div>
                </fieldset>
            </form>
        </div>
	</div>
	</div>

	<!-- #include file="includes/footer.inc" -->
</body>
</html>
